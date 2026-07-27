using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface ISemiFinishOpenReceivingDataAccess
    {
        Task<List<ItemViewModel>> GetArticlesAsync();
        Task<List<ItemProcessViewModel>> GetProcessesByArticleAsync(string itemId);
        Task<List<StoreLookupModel>> GetUserStoresAsync(string userName);
        Task<List<StoreShelfLookupModel>> GetStoreShelvesAsync(int storeRefId);
        Task<List<SemiFinishOpeningModel>> GetOpeningsByArticleAsync(string itemId);
        Task<string> GetNextLotNoAsync();
        Task<bool> CheckBatchNoExistsAsync(string batchNo);
        Task<int> AddOpeningAsync(CreateSemiFinishOpeningRequest request);
        Task UpdateRemarksAsync(int entryId, string remarks);
        Task<bool> DeleteOpeningAsync(int entryId);
    }
}
