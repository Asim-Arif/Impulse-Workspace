using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class SalarySheetService : ISalarySheetService
    {
        private readonly ISalarySheetDataAccess _dataAccess;
        private readonly IAuditService           _auditService;
        private readonly IVoucherService         _voucherService;

        public SalarySheetService(
            ISalarySheetDataAccess dataAccess,
            IAuditService auditService,
            IVoucherService voucherService)
        {
            _dataAccess     = dataAccess;
            _auditService   = auditService;
            _voucherService = voucherService;
        }

        public async Task<(List<SalarySheetRowModel> Rows, SalarySheetKpiSummaryModel Kpi, bool IsFinalized)> GetSalarySheetDataAsync(SalarySheetFilterDto filter)
        {
            bool isFinalized = await _dataAccess.IsSalarySheetFinalizedAsync(filter.Year, filter.Month, filter.IsExternal);

            List<SalarySheetRowModel> rows;
            if (isFinalized)
            {
                rows = await _dataAccess.GetFinalizedSalarySheetAsync(filter);
            }
            else
            {
                rows = await _dataAccess.GetUnfinalizedSalarySheetAsync(filter);
            }

            var kpi = new SalarySheetKpiSummaryModel
            {
                TotalEmployees    = rows.Count,
                GrossEarnedSalary = rows.Sum(r => r.GrossEarnings),
                TotalOvertime     = rows.Sum(r => r.OTAmount),
                TotalDeductions   = rows.Sum(r => r.TaxDeduction + r.ShortHoursDeduction +
                                                  r.ShortTermAdvanceDeduction + r.LongTermLoanDeduction +
                                                  r.UnionFundDeduction + r.FineDeduction + r.EOBIDeduction),
                NetPayableSalary  = rows.Sum(r => r.NetPayable)
            };

            return (rows, kpi, isFinalized);
        }

        public async Task FinalizeSalarySheetAsync(int year, int month, bool isExternal, List<SalarySheetRowModel> rows)
        {
            if (!rows.Any())
            {
                throw new InvalidOperationException("No salary rows available to finalize.");
            }

            string userName    = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();

            await _dataAccess.FinalizeSalarySheetAsync(year, month, isExternal, rows, userName, machineName);

            // Post centralized Journal Vouchers for the finalized salary sheet
            await AddSalaryVouchersAsync(year, month);
        }

        public async Task UpdateSalaryPaidStatusAsync(string empId, int year, int month, bool isPaid)
        {
            string userName    = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();

            await _dataAccess.UpdateSalaryPaidStatusAsync(empId, year, month, isPaid, userName, machineName);
        }

        public async Task UpdateLongTermDeductionOverrideAsync(string empId, DateTime monthDate, decimal amount)
        {
            string userName    = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();

            await _dataAccess.UpdateLongTermDeductionOverrideAsync(empId, monthDate, amount, userName, machineName);
        }

        public async Task TransferNegativeSalaryAsync(string empId, decimal amountToAdjust, DateTime monthDate)
        {
            await _dataAccess.TransferNegativeSalaryAsync(empId, amountToAdjust, monthDate);
        }

        public async Task<int> AddSalaryVouchersAsync(int year, int month)
        {
            var summaries = await _dataAccess.GetDepartmentSalaryVoucherSummariesAsync(year, month);
            if (summaries == null || summaries.Count == 0)
            {
                return 0;
            }

            var (stAccNo, ltAccNo) = await _dataAccess.GetSalaryVoucherGeneralAccountsAsync();

            DateTime dated = new DateTime(year, month, DateTime.DaysInMonth(year, month));
            if (dated.DayOfWeek == DayOfWeek.Sunday)
            {
                dated = dated.AddDays(-1);
            }

            int count = 0;
            foreach (var row in summaries)
            {
                if (row.TotalAmt <= 0)
                    continue;

                var voucher = new VoucherViewModel
                {
                    DT = dated,
                    Notes = $"Salary Of  [{row.DeptID}] {row.Name}",
                    JVType = 0,
                    PostedThroughJVForm = false,
                    LineItems = new List<VoucherLineItemViewModel>()
                };

                voucher.VchrNo = await _voucherService.GetNextVchrNo(dated, "JV");

                void AddLineItem(string accNo, string desc, decimal debit, decimal credit)
                {
                    if (debit == 0 && credit == 0) return;
                    if (string.IsNullOrWhiteSpace(accNo)) return;
                    voucher.LineItems.Add(new VoucherLineItemViewModel
                    {
                        AccNo       = accNo,
                        AccTitle    = string.Empty,
                        Description = desc,
                        Debit       = Math.Round(debit, 2),
                        Credit      = Math.Round(credit, 2)
                    });
                }

                string desc = voucher.Notes;

                // 1. Tax (Credit liability)
                if (row.TaxAmt > 0)
                {
                    AddLineItem(row.TaxAccNo, desc, 0, row.TaxAmt);
                }
                // 2. Salary Expense (Debit)
                if (row.TotalAmt > 0)
                {
                    AddLineItem(row.SalaryAccNo, desc, row.TotalAmt, 0);
                }
                // 3. Overtime Expense (Debit)
                if (row.OTAmt != 0)
                {
                    AddLineItem(row.AccruedUnionAccNo, desc, row.OTAmt, 0);
                }
                // 4. Bonus/Scholarship Expense (Debit)
                if (row.ZeroAbsentBonus != 0)
                {
                    AddLineItem(row.SchAccNo, desc, row.ZeroAbsentBonus, 0);
                }
                // 5. Short Term Advance Clearance (Credit)
                if (row.ShortAmt != 0)
                {
                    AddLineItem(stAccNo, desc, 0, row.ShortAmt);
                }
                // 6. Long Term Loan Clearance (Credit)
                if (row.LongAmt != 0)
                {
                    AddLineItem(ltAccNo, desc, 0, row.LongAmt);
                }
                // 7. EOBI Liability (Credit)
                if (row.EOBIAmt != 0)
                {
                    AddLineItem(row.AccruedEOBIAccNo, desc, 0, row.EOBIAmt);
                }
                // 8. One Percent / Fine
                if (row.DedOnePercentAmt < 0)
                {
                    AddLineItem(row.OnePercentAccNo, desc, 0, Math.Abs(row.DedOnePercentAmt));
                }
                else if (row.DedOnePercentAmt > 0)
                {
                    AddLineItem(row.OnePercentAccNo, desc, row.DedOnePercentAmt, 0);
                }

                // 9. Accrued Salaries Payable (Credit Liability to balance total debits/credits)
                decimal totalDebits  = voucher.LineItems.Sum(x => x.Debit);
                decimal totalCredits = voucher.LineItems.Sum(x => x.Credit);
                decimal accruedPayableAmt = totalDebits - totalCredits;
                if (accruedPayableAmt > 0)
                {
                    AddLineItem(row.AccruedAccNo, desc, 0, accruedPayableAmt);
                }
                else if (accruedPayableAmt < 0)
                {
                    AddLineItem(row.AccruedAccNo, desc, Math.Abs(accruedPayableAmt), 0);
                }

                await _voucherService.SaveVoucherAsync(voucher);
                count++;
            }

            return count;
        }
    }
}
