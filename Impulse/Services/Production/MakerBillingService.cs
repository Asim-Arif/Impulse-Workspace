using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerBillingService : IMakerBillingService
    {
        private readonly IMakerBillingDataAccess _dataAccess;
        private readonly IVoucherService _voucherService;
        private readonly IAuditService _auditService;

        public MakerBillingService(
            IMakerBillingDataAccess dataAccess,
            IVoucherService voucherService,
            IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _voucherService = voucherService;
            _auditService = auditService;
        }

        public async Task<MakerBillingLookups> GetLookupsAsync()
        {
            var makers = await _dataAccess.GetMakersAsync();
            var cats = await _dataAccess.GetCategoriesAsync();
            var debits = await _dataAccess.GetDebitHeadsAsync();

            return new MakerBillingLookups
            {
                Makers = makers,
                Categories = cats,
                DebitHeads = debits
            };
        }

        public Task<List<MakerBillingLotItem>> GetMakerBillingLotsAsync(MakerBillingFilter filter) => _dataAccess.GetMakerBillingLotsAsync(filter);
        public Task<decimal> GetShortTermDeductionAsync(int vendId) => _dataAccess.GetShortTermDeductionAsync(vendId);
        public Task<decimal> GetLongTermBalanceAsync(int vendId) => _dataAccess.GetLongTermBalanceAsync(vendId);
        public Task<List<MakerRepairItem>> GetRepairWastageItemsAsync(MakerBillingFilter filter) => _dataAccess.GetRepairWastageItemsAsync(filter);
        public Task<List<MakerMaterialDedItem>> GetMaterialDeductionsAsync(int vendId, string masterPoNo) => _dataAccess.GetMaterialDeductionsAsync(vendId, masterPoNo);
        public Task<List<MakerHiddenLotItem>> GetHiddenLotsAsync(int vendId) => _dataAccess.GetHiddenLotsAsync(vendId);
        public Task<List<MakerHiddenRMItem>> GetHiddenRMItemsAsync(int vendId) => _dataAccess.GetHiddenRMItemsAsync(vendId);

        public Task UpdateLotRateAsync(long refId, decimal rate) => _dataAccess.UpdateLotRateAsync(refId, rate);
        public Task UpdateMaterialPriceAsync(long entryId, decimal price) => _dataAccess.UpdateMaterialPriceAsync(entryId, price);
        public Task MarkLotsAsAuditedAsync(List<long> vrdEntryIds, bool audited) => _dataAccess.MarkLotsAsAuditedAsync(vrdEntryIds, audited);

        public Task HideLotFromBillingAsync(long vrdEntryId)
        {
            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();
            return _dataAccess.HideLotFromBillingAsync(vrdEntryId, userName, machineName);
        }

        public Task RestoreHiddenLotAsync(long hiddenEntryId) => _dataAccess.RestoreHiddenLotAsync(hiddenEntryId);

        public Task HideRMFromBillingAsync(long rmidRefId)
        {
            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();
            return _dataAccess.HideRMFromBillingAsync(rmidRefId, userName, machineName);
        }

        public Task RestoreHiddenRMAsync(long rmidRefId) => _dataAccess.RestoreHiddenRMAsync(rmidRefId);

        public Task<string?> GetItemPicBase64Async(string itemCode) => _dataAccess.GetItemPicBase64Async(itemCode);

        public async Task<PostMakerBillResult> PostMakerBillAsync(PostMakerBillRequest request)
        {
            if (request.VendID <= 0)
            {
                return new PostMakerBillResult { Success = false, Message = "Please select a valid Maker." };
            }

            if (request.SelectedLots == null || !request.SelectedLots.Any())
            {
                return new PostMakerBillResult { Success = false, Message = "No orders/lots selected for billing." };
            }

            if (string.IsNullOrWhiteSpace(request.DebitHeadAccNo))
            {
                return new PostMakerBillResult { Success = false, Message = "Please select a Debit Head account." };
            }

            decimal totalValue = request.SelectedLots.Sum(x => x.Value);
            decimal repairRelax = request.RejectionRelaxation ? (totalValue / 100m) * 2m : 0;
            decimal repairAmtToDeduct = Math.Max(0, request.RepairAmount - repairRelax);

            decimal dAmtToPost = totalValue - request.ManualDeductionAmount - request.MaterialDeductionAmount - repairAmtToDeduct;

            if (dAmtToPost <= 0)
            {
                return new PostMakerBillResult { Success = false, Message = "Invalid Amount to Post. Deductions exceed or equal total bill value." };
            }

            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();
            request.UserName = userName;
            request.MachineName = machineName;

            // 1. Centralized Vouchering: Generate JV Voucher Number
            string vchrNo = await _voucherService.GetNextVchrNo(request.PostingDate, "JV");
            string makerAccNo = await _dataAccess.GetMakerAccountNoAsync(request.VendID);
            string makerCode = await _dataAccess.GetMakerCodeAsync(request.VendID);
            long nextBillNo = await _dataAccess.GetNextBillNoAsync(request.VendID);

            string desc = $"Computer Bill No. {makerCode} / {nextBillNo} - {request.PostingDate.Year}";

            var voucher = new VoucherViewModel
            {
                VchrNo = vchrNo,
                DT = request.PostingDate,
                UserName = userName,
                MachineName = machineName,
                Notes = desc,
                LineItems = new List<VoucherLineItemViewModel>
                {
                    new VoucherLineItemViewModel
                    {
                        AccNo = request.DebitHeadAccNo,
                        Description = desc,
                        Debit = dAmtToPost,
                        Credit = 0
                    },
                    new VoucherLineItemViewModel
                    {
                        AccNo = makerAccNo,
                        Description = desc,
                        Debit = 0,
                        Credit = dAmtToPost
                    }
                }
            };

            // Post via centralized VoucherService
            await _voucherService.SaveVoucherAsync(voucher);

            // 2. Commit MakerPostedBills database transaction
            return await _dataAccess.SaveMakerPostedBillTransactionAsync(request, vchrNo);
        }
    }
}
