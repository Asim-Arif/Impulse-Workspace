using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface ITransferToReadyFinishStockDataAccess
    {
        Task<List<LookupItemInt>> GetMakersAsync();
        Task<List<LookupItemString>> GetArticlesAsync();
        Task<List<StoreLookupDto>> GetStoresAsync();
        Task<List<ShelfLookupDto>> GetShelvesByStoreAsync(int storeId);
        Task<string> GetShelfRemarksAsync(string itemId, int shelfRefId);
        Task<List<ReadyFinishStockAvailableItem>> GetAvailableLotsAsync(ReadyFinishStockFilter filter);
        Task<PostTransferResult> PostTransferToReadyFinishAsync(PostTransferRequest request);
    }
}
