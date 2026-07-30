using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class ClearShortTermService : IClearShortTermService
    {
        private readonly IClearShortTermDataAccess _dataAccess;
        private readonly IVoucherService           _voucherService;
        private readonly IDBHelperService          _dbHelperService;
        private readonly IAuditService             _auditService;

        public ClearShortTermService(
            IClearShortTermDataAccess dataAccess,
            IVoucherService           voucherService,
            IDBHelperService          dbHelperService,
            IAuditService             auditService)
        {
            _dataAccess      = dataAccess;
            _voucherService  = voucherService;
            _dbHelperService = dbHelperService;
            _auditService    = auditService;
        }

        // ── Public API ───────────────────────────────────────────────────────

        public async Task<List<ShortTermAdvanceItem>> GetOpenAdvancesAsync(string empId)
        {
            return await _dataAccess.GetOpenAdvancesAsync(empId);
        }

        public async Task<string> GetEmpShortTermAccNoAsync()
        {
            return await _dbHelperService.getSingleStringValue(
                "DataValue", "GeneralData", "WHERE DataName = 'EmpShortTermAccNo'");
        }

        /// <summary>
        /// Core save:
        ///  1. Resolve EmpShortTermAccNo from GeneralData
        ///  2. Unless "Don't Post Voucher": build a two-line VoucherViewModel
        ///     (debit receiving account, credit employee short-term advance
        ///     account, for the SUM of the selected advance amounts) and post
        ///     it via the centralized IVoucherService (CRV for Cash In Hand,
        ///     JV for Other Account)
        ///  3. Record the clearance header + detail rows (+ Advances.Cleared=2)
        ///     via DAC in one SQL transaction
        ///  4. Return the posted VchrNo (empty when no voucher was posted)
        /// </summary>
        public async Task<string> SaveAsync(ClearShortTermDto dto)
        {
            // 1. Resolve short-term advance account
            string empSTAccNo = await GetEmpShortTermAccNoAsync();
            if (string.IsNullOrWhiteSpace(empSTAccNo))
                throw new InvalidOperationException(
                    "Employee Short-Term Advance Account (EmpShortTermAccNo) is not configured in General Data. " +
                    "Please contact your system administrator.");

            decimal totalAmount = dto.SelectedAdvances.Sum(a => a.Amount);

            string vchrNo = string.Empty;

            // 2. Post voucher unless suppressed
            if (!dto.DontPostVoucher)
            {
                string vchrType = dto.ReceiptMode == ClearAdvanceReceiptMode.Cash ? "CRV" : "JV";
                string theDesc  = $"Temp. Advance Received From [ {dto.EmpID} ] {dto.EmpDisplayName}";

                var voucher = new VoucherViewModel
                {
                    VchrNo      = await _voucherService.GetNextVchrNo(dto.ClearDate, vchrType),
                    DT          = dto.ClearDate,
                    UserName    = _auditService.GetCurrentUserName(),
                    MachineName = _auditService.GetClientIpAddress(),
                    Notes       = dto.Remarks,
                    LineItems   = new List<VoucherLineItemViewModel>
                    {
                        // Debit side: receiving account (cash in hand / other account)
                        new VoucherLineItemViewModel
                        {
                            AccNo       = dto.ReceiptAccNo,
                            Description = theDesc,
                            Debit       = totalAmount,
                            Credit      = 0
                        },
                        // Credit side: employee short-term advance account
                        new VoucherLineItemViewModel
                        {
                            AccNo       = empSTAccNo,
                            Description = theDesc,
                            Debit       = 0,
                            Credit      = totalAmount
                        }
                    }
                };

                vchrNo = await _voucherService.SaveVoucherAsync(voucher);
            }

            // 3. Record the clearance
            await _dataAccess.SaveClearanceAsync(
                dto, vchrNo,
                _auditService.GetCurrentUserName(),
                _auditService.GetClientIpAddress());

            return vchrNo;
        }
    }
}
