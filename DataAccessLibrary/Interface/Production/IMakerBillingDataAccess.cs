using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IMakerBillingDataAccess
    {
        Task<List<LookupItemInt>> GetMakersAsync();
        Task<List<LookupItemInt>> GetCategoriesAsync();
        Task<List<LookupItemString>> GetDebitHeadsAsync();
        Task<string> GetMakerAccountNoAsync(int vendId);
        Task<string> GetMakerCodeAsync(int vendId);
        Task<long> GetNextBillNoAsync(int vendId);

        Task<List<MakerBillingLotItem>> GetMakerBillingLotsAsync(MakerBillingFilter filter);
        Task<decimal> GetShortTermDeductionAsync(int vendId);
        Task<decimal> GetLongTermBalanceAsync(int vendId);
        Task<List<MakerRepairItem>> GetRepairWastageItemsAsync(MakerBillingFilter filter);
        Task<List<MakerMaterialDedItem>> GetMaterialDeductionsAsync(int vendId, string masterPoNo);
        Task<List<MakerHiddenLotItem>> GetHiddenLotsAsync(int vendId);
        Task<List<MakerHiddenRMItem>> GetHiddenRMItemsAsync(int vendId);

        Task UpdateLotRateAsync(long refId, decimal rate);
        Task UpdateMaterialPriceAsync(long entryId, decimal price);
        Task MarkLotsAsAuditedAsync(List<long> vrdEntryIds, bool audited);

        Task HideLotFromBillingAsync(long vrdEntryId, string userName, string machineName);
        Task RestoreHiddenLotAsync(long hiddenEntryId);

        Task HideRMFromBillingAsync(long rmidRefId, string userName, string machineName);
        Task RestoreHiddenRMAsync(long rmidRefId);

        Task<string?> GetItemPicBase64Async(string itemCode);

        Task<PostMakerBillResult> SaveMakerPostedBillTransactionAsync(PostMakerBillRequest request, string voucherNo);
    }
}
