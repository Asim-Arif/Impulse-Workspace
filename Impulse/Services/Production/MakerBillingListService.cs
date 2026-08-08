using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public class MakerBillingListService : IMakerBillingListService
    {
        private readonly IMakerBillingListDataAccess _dataAccess;
        private readonly IVoucherService _voucherService;
        private readonly IAuditService _auditService;

        public MakerBillingListService(
            IMakerBillingListDataAccess dataAccess,
            IVoucherService voucherService,
            IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _voucherService = voucherService;
            _auditService = auditService;
        }

        public Task<List<LookupItemInt>> GetMakersAsync() => _dataAccess.GetMakersAsync();

        public Task<List<MakerBillingListItem>> GetPostedBillsListAsync(MakerBillingListFilter filter) => _dataAccess.GetPostedBillsListAsync(filter);

        public Task<List<LookupItemString>> GetPaymentCreditAccountsAsync() => _dataAccess.GetPaymentCreditAccountsAsync();

        public Task<string?> GetMakerAccountNoAsync(int vendId) => _dataAccess.GetMakerAccountNoAsync(vendId);

        public async Task<string> PostPaymentVoucherAsync(MakerBillingPaymentVoucherRequest request)
        {
            if (request.Amount <= 0)
                throw new InvalidOperationException("Payment amount must be greater than zero.");

            string makerAccNo = await _dataAccess.GetMakerAccountNoAsync(request.VendID)
                ?? request.DebitAccountNo;

            if (string.IsNullOrWhiteSpace(makerAccNo))
                throw new InvalidOperationException("Maker does not have an Account Number configured.");

            if (string.IsNullOrWhiteSpace(request.CreditAccountNo))
                throw new InvalidOperationException("Please select a valid Cash / Bank Credit Account.");

            string vchrType = request.PaymentMode == "Bank" ? "BPV" : "CPV";
            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();

            string vchrNo = await _voucherService.GetNextVchrNo(request.PaymentDate, vchrType);

            string description = !string.IsNullOrWhiteSpace(request.Remarks)
                ? request.Remarks.Trim()
                : $"Payment against Maker Bill #{request.BillNo} (Vchr #{request.BillVoucherNo}) to {request.MakerName}";

            var voucher = new VoucherViewModel
            {
                VchrNo = vchrNo,
                DT = request.PaymentDate,
                UserName = userName,
                MachineName = machineName,
                Notes = description,
                LineItems = new List<VoucherLineItemViewModel>
                {
                    // Debit Maker
                    new VoucherLineItemViewModel
                    {
                        AccNo = makerAccNo,
                        Debit = request.Amount,
                        Credit = 0,
                        Description = description
                    },
                    // Credit Cash / Bank
                    new VoucherLineItemViewModel
                    {
                        AccNo = request.CreditAccountNo,
                        Debit = 0,
                        Credit = request.Amount,
                        Description = description
                    }
                }
            };

            await _voucherService.SaveVoucherAsync(voucher);
            await _dataAccess.RecordPaymentVoucherAsync(request.MPB_RefID, vchrNo, request.Amount);
            return vchrNo;
        }

        public Task<bool> DeletePostedBillAsync(long entryId, string userName) => _dataAccess.DeletePostedBillAsync(entryId, userName);
    }
}
