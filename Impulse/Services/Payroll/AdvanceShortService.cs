using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class AdvanceShortService : IAdvanceShortService
    {
        private readonly IAdvanceShortDataAccess _dataAccess;
        private readonly IVoucherService         _voucherService;
        private readonly IDBHelperService        _dbHelperService;
        private readonly IAuditService           _auditService;

        public AdvanceShortService(
            IAdvanceShortDataAccess dataAccess,
            IVoucherService         voucherService,
            IDBHelperService        dbHelperService,
            IAuditService           auditService)
        {
            _dataAccess      = dataAccess;
            _voucherService  = voucherService;
            _dbHelperService = dbHelperService;
            _auditService    = auditService;
        }

        // ── Public API ───────────────────────────────────────────────────────

        public async Task<AdvanceShortSummaryModel> GetEmployeeSummaryAsync(string empId)
        {
            return await _dataAccess.GetEmployeeAdvanceSummaryAsync(empId);
        }

        public async Task<string> GetEmpShortTermAccNoAsync()
        {
            return await _dbHelperService.getSingleStringValue(
                "DataValue", "GeneralData", "WHERE DataName = 'EmpShortTermAccNo'");
        }

        public async Task<string> GetPreviewVchrNoAsync(DateTime date, AdvanceShortPaymentMode mode)
        {
            string vchrType = MapVchrType(mode);
            return await _voucherService.GetNextVchrNo(date, vchrType);
        }

        /// <summary>
        /// Core save:
        ///  1. Resolve EmpShortTermAccNo from GeneralData
        ///  2. Generate VchrNo using GetPreviewVchrNoAsync
        ///  3. Build VoucherViewModel with two lines (credit payment side, debit short-term loan side)
        ///  4. For BPV: populate ChequeDetails (including CDate and AccNo)
        ///  5. Call centralized SaveVoucherAsync
        ///  6. INSERT into Advances table (Type=0) via DAC
        ///  7. Return the posted VchrNo
        /// </summary>
        public async Task<string> SaveAsync(AdvanceShortDto dto)
        {
            // 1. Resolve short-term loan account
            string empSTAccNo = await GetEmpShortTermAccNoAsync();
            if (string.IsNullOrWhiteSpace(empSTAccNo))
                throw new InvalidOperationException(
                    "Employee Short-Term Loan Account (EmpShortTermAccNo) is not configured in General Data. " +
                    "Please contact your system administrator.");

            // 2. Build description and preview VchrNo
            string theDesc = string.IsNullOrWhiteSpace(dto.Description)
                ? $"Short Term Advance Paid — Amount: {dto.Amount:N2}"
                : dto.Description;

            string vchrNo = await GetPreviewVchrNoAsync(dto.IssueDate, dto.PaymentMode);

            // 3. Build voucher
            var voucher = new VoucherViewModel
            {
                VchrNo      = vchrNo,
                DT          = dto.IssueDate,
                UserName    = _auditService.GetCurrentUserName(),
                MachineName = _auditService.GetClientIpAddress(),
                Notes       = theDesc,
                LineItems   = new List<VoucherLineItemViewModel>
                {
                    // Credit side: payment account (cash / other acc / bank)
                    new VoucherLineItemViewModel
                    {
                        AccNo       = dto.PaymentAccNo,
                        Description = theDesc,
                        Debit       = 0,
                        Credit      = dto.Amount
                    },
                    // Debit side: employee short-term loan account
                    new VoucherLineItemViewModel
                    {
                        AccNo       = empSTAccNo,
                        Description = theDesc,
                        Debit       = dto.Amount,
                        Credit      = 0
                    }
                }
            };

            // 4. For BPV: add cheque details so the centralized service creates the Cheque record
            if (dto.PaymentMode == AdvanceShortPaymentMode.Bank)
            {
                voucher.ChequeDetails = new ChequeViewModel
                {
                    CDate       = dto.IssueDate,
                    AccNo       = dto.PaymentAccNo,
                    BankAccNo   = dto.PaymentAccNo,
                    BankID      = dto.BankID,
                    ChequeNo    = dto.ChequeNo,
                    ChqBookNo   = dto.ChqBookNo,
                    chequeType  = dto.ChequeType,
                    ChequeDate  = dto.ChequeDate,
                    Amount      = dto.Amount,
                    Description = theDesc,
                    Payment     = true,
                    Posted      = false,
                    Bounced     = false
                };
            }

            // 5. Post voucher through centralized service
            string postedVchrNo = await _voucherService.SaveVoucherAsync(voucher);

            // 6. Record in Advances table (Type = 0)
            await _dataAccess.SaveAdvanceShortRecordAsync(dto, postedVchrNo);

            return postedVchrNo;
        }

        // ── Helpers ──────────────────────────────────────────────────────────

        private static string MapVchrType(AdvanceShortPaymentMode mode) => mode switch
        {
            AdvanceShortPaymentMode.Cash         => "CPV",
            AdvanceShortPaymentMode.OtherAccount => "JV",
            AdvanceShortPaymentMode.Bank         => "BPV",
            _                                    => "CPV"
        };
    }
}
