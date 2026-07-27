using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class SemiFinishOpenReceivingService : ISemiFinishOpenReceivingService
    {
        private readonly ISemiFinishOpenReceivingDataAccess _dataAccess;

        public SemiFinishOpenReceivingService(ISemiFinishOpenReceivingDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ItemViewModel>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();

        public Task<List<ItemProcessViewModel>> GetProcessesByArticleAsync(string itemId) => _dataAccess.GetProcessesByArticleAsync(itemId);

        public Task<List<StoreLookupModel>> GetUserStoresAsync(string userName) => _dataAccess.GetUserStoresAsync(userName);

        public Task<List<StoreShelfLookupModel>> GetStoreShelvesAsync(int storeRefId) => _dataAccess.GetStoreShelvesAsync(storeRefId);

        public Task<List<SemiFinishOpeningModel>> GetOpeningsByArticleAsync(string itemId) => _dataAccess.GetOpeningsByArticleAsync(itemId);

        public Task<string> GetNextLotNoAsync() => _dataAccess.GetNextLotNoAsync();

        public Task<bool> CheckBatchNoExistsAsync(string batchNo) => _dataAccess.CheckBatchNoExistsAsync(batchNo);

        public Task<int> AddOpeningAsync(CreateSemiFinishOpeningRequest request) => _dataAccess.AddOpeningAsync(request);

        public Task UpdateRemarksAsync(int entryId, string remarks) => _dataAccess.UpdateRemarksAsync(entryId, remarks);

        public Task<bool> DeleteOpeningAsync(int entryId) => _dataAccess.DeleteOpeningAsync(entryId);
    }
}
