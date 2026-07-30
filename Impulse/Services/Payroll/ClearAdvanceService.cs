using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class ClearAdvanceService : IClearAdvanceService
    {
        private readonly IClearAdvanceDataAccess _dataAccess;
        private readonly IVoucherService         _voucherService;
        private readonly IDBHelperService        _dbHelperService;
        private readonly IAuditService           _auditService;

        public ClearAdvanceService(
            IClearAdvanceDataAccess dataAccess,
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

        public async Task<ClearAdvanceSummaryModel> GetSummaryAsync(string empId)
        {
            return await _dataAccess.GetSummaryAsync(empId);
        }

        public async Task<string> GetEmpLongTermAccNoAsync()
        {
            return await _dbHelperService.getSingleStringValue(
                "DataValue", "GeneralData", "WHERE DataName = 'EmpLongTermAccNo'");
        }

        /// <summary>
        /// Core save:
        ///  1. Resolve EmpLongTermAccNo from GeneralData
        ///  2. Unless "Don't Post Voucher": build a two-line VoucherViewModel
        ///     (debit receiving account, credit employee long-term loan account)
        ///     and post it via the centralized IVoucherService (CRV for Cash In
        ///     Hand, JV for Other Account)
        ///  3. INSERT into AmtCleared (+ conditional Advances.DAmount update) via DAC
        ///  4. Return the posted VchrNo (empty when no voucher was posted)
        /// </summary>
        public async Task<string> SaveAsync(ClearAdvanceDto dto)
        {
            // 1. Resolve loan-side account
            string empLTAccNo = await GetEmpLongTermAccNoAsync();
            if (string.IsNullOrWhiteSpace(empLTAccNo))
                throw new InvalidOperationException(
                    "Employee Long-Term Loan Account (EmpLongTermAccNo) is not configured in General Data. " +
                    "Please contact your system administrator.");

            string vchrNo = string.Empty;

            // 2. Post voucher unless suppressed
            if (!dto.DontPostVoucher)
            {
                string vchrType = dto.ReceiptMode == ClearAdvanceReceiptMode.Cash ? "CRV" : "JV";
                string theDesc  = $"Permanent Advance Received From [ {dto.EmpID} ] {dto.EmpDisplayName}";

                var voucher = new VoucherViewModel
                {
                    VchrNo      = await _voucherService.GetNextVchrNo(dto.ClearDate, vchrType),
                    DT          = dto.ClearDate,
                    UserName    = _auditService.GetCurrentUserName(),
                    MachineName = _auditService.GetClientIpAddress(),
                    Notes       = dto.Description,
                    LineItems   = new List<VoucherLineItemViewModel>
                    {
                        // Debit side: receiving account (cash in hand / other account)
                        new VoucherLineItemViewModel
                        {
                            AccNo       = dto.ReceiptAccNo,
                            Description = theDesc,
                            Debit       = dto.AmountReceived,
                            Credit      = 0
                        },
                        // Credit side: employee long-term loan account
                        new VoucherLineItemViewModel
                        {
                            AccNo       = empLTAccNo,
                            Description = theDesc,
                            Debit       = 0,
                            Credit      = dto.AmountReceived
                        }
                    }
                };

                vchrNo = await _voucherService.SaveVoucherAsync(voucher);
            }

            // 3. Record the clearance
            await _dataAccess.SaveClearanceAsync(dto, vchrNo);

            return vchrNo;
        }
    }
}
