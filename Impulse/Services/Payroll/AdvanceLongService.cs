using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using System;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class AdvanceLongService : IAdvanceLongService
    {
        private readonly IAdvanceLongDataAccess _dataAccess;
        private readonly IVoucherService        _voucherService;
        private readonly IDBHelperService       _dbHelperService;
        private readonly IAuditService          _auditService;

        public AdvanceLongService(
            IAdvanceLongDataAccess dataAccess,
            IVoucherService        voucherService,
            IDBHelperService       dbHelperService,
            IAuditService          auditService)
        {
            _dataAccess      = dataAccess;
            _voucherService  = voucherService;
            _dbHelperService = dbHelperService;
            _auditService    = auditService;
        }

        // ── Public API ───────────────────────────────────────────────────────

        public async Task<AdvanceLongSummaryModel> GetEmployeeSummaryAsync(string empId)
        {
            return await _dataAccess.GetEmployeeAdvanceSummaryAsync(empId);
        }

        public async Task<string> GetEmpLongTermAccNoAsync()
        {
            return await _dbHelperService.getSingleStringValue(
                "DataValue", "GeneralData", "WHERE DataName = 'EmpLongTermAccNo'");
        }

        public async Task<string> GetPreviewVchrNoAsync(DateTime date, AdvanceLongPaymentMode mode)
        {
            string vchrType = MapVchrType(mode);
            return await _voucherService.GetNextVchrNo(date, vchrType);
        }

        /// <summary>
        /// Core save:
        ///  1. Resolve EmpLongTermAccNo from GeneralData
        ///  2. Build VoucherViewModel with two lines (credit payment side, debit loan side)
        ///  3. For BPV: populate ChequeDetails
        ///  4. Call centralized SaveVoucherAsync
        ///  5. INSERT into Advances table via DAC
        ///  6. Return the posted VchrNo
        /// </summary>
        public async Task<string> SaveAsync(AdvanceLongDto dto)
        {
            // 1. Resolve loan-side account
            string empLTAccNo = await GetEmpLongTermAccNoAsync();
            if (string.IsNullOrWhiteSpace(empLTAccNo))
                throw new InvalidOperationException(
                    "Employee Long-Term Loan Account (EmpLongTermAccNo) is not configured in General Data. " +
                    "Please contact your system administrator.");

            // 2. Build description and get next Voucher Number
            string theDesc = $"Long Term Advance Paid — Amount: {dto.Amount:N2}  Monthly Deduction: {dto.DAmount:N2}";
            string vchrNo  = await GetPreviewVchrNoAsync(dto.IssueDate, dto.PaymentMode);

            // 3. Build voucher
            var voucher = new VoucherViewModel
            {
                VchrNo     = vchrNo,
                DT         = dto.IssueDate,
                UserName   = _auditService.GetCurrentUserName(),
                MachineName= _auditService.GetClientIpAddress(),
                Notes      = dto.Description,
                LineItems  = new System.Collections.Generic.List<VoucherLineItemViewModel>
                {
                    // Credit side: payment account (cash / other acc / bank)
                    new VoucherLineItemViewModel
                    {
                        AccNo       = dto.PaymentAccNo,
                        Description = theDesc,
                        Debit       = 0,
                        Credit      = dto.Amount
                    },
                    // Debit side: employee long-term loan account
                    new VoucherLineItemViewModel
                    {
                        AccNo       = empLTAccNo,
                        Description = theDesc,
                        Debit       = dto.Amount,
                        Credit      = 0
                    }
                }
            };

            // 4. For BPV: add cheque details so the centralized service creates the Cheque record
            if (dto.PaymentMode == AdvanceLongPaymentMode.Bank)
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

            // 6. Record in Advances table
            await _dataAccess.SaveAdvanceLongRecordAsync(dto, postedVchrNo);

            return postedVchrNo;
        }

        // ── Helpers ──────────────────────────────────────────────────────────

        private static string MapVchrType(AdvanceLongPaymentMode mode) => mode switch
        {
            AdvanceLongPaymentMode.Cash         => "CPV",
            AdvanceLongPaymentMode.OtherAccount => "JV",
            AdvanceLongPaymentMode.Bank         => "BPV",
            _                                   => "CPV"
        };
    }
}
