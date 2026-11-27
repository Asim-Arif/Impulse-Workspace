using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Setup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerBillingService : IMakerBillingService
    {
        private readonly IMakerBillingDataAccess _dataAccess;
        private readonly IAuditService _auditService;
        private readonly IUserPermissionService _permissionService;

        public MakerBillingService(
            IMakerBillingDataAccess dataAccess,
            IAuditService auditService,
            IUserPermissionService permissionService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
            _permissionService = permissionService;
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
        public Task<MakerLongTermLoanInfo> GetLongTermLoanInfoAsync(int vendId) => _dataAccess.GetLongTermLoanInfoAsync(vendId);
        public Task<List<MakerRepairItem>> GetRepairWastageItemsAsync(MakerBillingFilter filter) => _dataAccess.GetRepairWastageItemsAsync(filter);
        public Task<List<MakerMaterialDedItem>> GetMaterialDeductionsAsync(int vendId, string masterPoNo) => _dataAccess.GetMaterialDeductionsAsync(vendId, masterPoNo);
        public Task<List<MakerHiddenLotItem>> GetHiddenLotsAsync(int vendId) => _dataAccess.GetHiddenLotsAsync(vendId);
        public Task<List<MakerHiddenRMItem>> GetHiddenRMItemsAsync(int vendId) => _dataAccess.GetHiddenRMItemsAsync(vendId);

        public async Task UpdateLotRateAsync(long refId, decimal rate)
        {
            await _permissionService.InitializeAsync();
            if (!_permissionService.CanEditMakerBillRate)
            {
                throw new UnauthorizedAccessException("You do not have permission to edit rates in Maker Billing.");
            }
            await _dataAccess.UpdateLotRateAsync(refId, rate);
        }

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
            await _permissionService.InitializeAsync();
            if (!_permissionService.CanPostMakerBill)
            {
                return new PostMakerBillResult { Success = false, Message = "You do not have permission to post Maker Bills." };
            }

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

            if (request.ShortTermDeduction < 0 || request.LongTermDeduction < 0)
            {
                return new PostMakerBillResult { Success = false, Message = "Loan deduction amounts cannot be negative." };
            }

            decimal maxShortTerm = await _dataAccess.GetShortTermDeductionAsync(request.VendID);
            if (request.ShortTermDeduction > maxShortTerm)
            {
                return new PostMakerBillResult { Success = false, Message = $"Invalid Short-Term Deduction. Max allowed is {maxShortTerm:N2}." };
            }

            var ltInfo = await _dataAccess.GetLongTermLoanInfoAsync(request.VendID);
            if (request.LongTermDeduction > ltInfo.RemainingBalance)
            {
                return new PostMakerBillResult { Success = false, Message = $"Invalid Long-Term Deduction. Max allowed is {ltInfo.RemainingBalance:N2}." };
            }

            decimal totalValue = Math.Round(request.SelectedLots.Sum(x => x.Value), 0, MidpointRounding.AwayFromZero);
            decimal repairRelax = request.RejectionRelaxation
                ? Math.Round((totalValue / 100m) * 2m, 0, MidpointRounding.AwayFromZero)
                : 0m;
            decimal repairAmtToDeduct = Math.Max(0m, Math.Round(request.RepairAmount, 0, MidpointRounding.AwayFromZero) - repairRelax);
            decimal dAmtToPost = totalValue
                - Math.Round(request.ManualDeductionAmount, 0, MidpointRounding.AwayFromZero)
                - Math.Round(request.MaterialDeductionAmount, 0, MidpointRounding.AwayFromZero)
                - repairAmtToDeduct;

            if (dAmtToPost <= 0)
            {
                return new PostMakerBillResult { Success = false, Message = "Invalid Amount to Post. Deductions exceed or equal total bill value." };
            }

            string makerAccNo = await _dataAccess.GetMakerAccountNoAsync(request.VendID);
            if (string.IsNullOrWhiteSpace(makerAccNo))
            {
                return new PostMakerBillResult { Success = false, Message = "Selected maker does not have a linked Account No. configured." };
            }

            string makerCode = await _dataAccess.GetMakerCodeAsync(request.VendID);
            request.MakerAccNo = makerAccNo;
            request.UserName = _auditService.GetCurrentUserName();
            request.MachineName = _auditService.GetClientIpAddress();

            // Atomic JV + MakerPostedBills + Repair/Wastage + ST/LT Loan clearing transaction
            return await _dataAccess.SaveMakerPostedBillTransactionAsync(request, makerCode);
        }
    }
}
