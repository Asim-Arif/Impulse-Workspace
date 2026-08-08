using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class TransferToReadyFinishStockService : ITransferToReadyFinishStockService
    {
        private readonly ITransferToReadyFinishStockDataAccess _dataAccess;

        public TransferToReadyFinishStockService(ITransferToReadyFinishStockDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<LookupItemInt>> GetMakersAsync() => _dataAccess.GetMakersAsync();
        public Task<List<LookupItemString>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();
        public Task<List<StoreLookupDto>> GetStoresAsync() => _dataAccess.GetStoresAsync();
        public Task<List<ShelfLookupDto>> GetShelvesByStoreAsync(int storeId) => _dataAccess.GetShelvesByStoreAsync(storeId);
        public Task<string> GetShelfRemarksAsync(string itemId, int shelfRefId) => _dataAccess.GetShelfRemarksAsync(itemId, shelfRefId);
        public Task<List<ReadyFinishStockAvailableItem>> GetAvailableLotsAsync(ReadyFinishStockFilter filter) => _dataAccess.GetAvailableLotsAsync(filter);
        public Task<PostTransferResult> PostTransferToReadyFinishAsync(PostTransferRequest request) => _dataAccess.PostTransferToReadyFinishAsync(request);
    }
}
