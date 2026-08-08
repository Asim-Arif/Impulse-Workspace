using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IMakerBillingService
    {
        Task<MakerBillingLookups> GetLookupsAsync();
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

        Task HideLotFromBillingAsync(long vrdEntryId);
        Task RestoreHiddenLotAsync(long hiddenEntryId);

        Task HideRMFromBillingAsync(long rmidRefId);
        Task RestoreHiddenRMAsync(long rmidRefId);

        Task<string?> GetItemPicBase64Async(string itemCode);

        Task<PostMakerBillResult> PostMakerBillAsync(PostMakerBillRequest request);
    }
}
