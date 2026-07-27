using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class SFMovementService : ISFMovementService
    {
        private readonly ISFMovementDataAccess _dataAccess;

        public SFMovementService(ISFMovementDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ItemViewModel>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();

        public Task<List<ItemProcessViewModel>> GetProcessesByArticleAsync(string itemId) => _dataAccess.GetProcessesByArticleAsync(itemId);

        public Task<List<SFTargetStoreModel>> GetFromStoresAsync(string itemId, int processId) => _dataAccess.GetFromStoresAsync(itemId, processId);

        public Task<List<SFFromLocationModel>> GetFromLocationsAsync(string itemId, int processId, int storeRefId) => _dataAccess.GetFromLocationsAsync(itemId, processId, storeRefId);

        public Task<List<SFTargetStoreModel>> GetTargetStoresAsync(string userName) => _dataAccess.GetTargetStoresAsync(userName);

        public Task<List<SFTargetRackModel>> GetTargetRacksAsync(int storeRefId) => _dataAccess.GetTargetRacksAsync(storeRefId);

        public Task<List<SFTargetShelfModel>> GetTargetShelvesAsync(int storeRefId, string rackNo) => _dataAccess.GetTargetShelvesAsync(storeRefId, rackNo);

        public Task<int> GetTargetLocationQtyAsync(string itemId, int processId, int shelfRefId) => _dataAccess.GetTargetLocationQtyAsync(itemId, processId, shelfRefId);

        public Task<bool> TransferLocationAsync(SFMovementRequest request) => _dataAccess.TransferLocationAsync(request);
    }
}
