using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class FinishMovementService : IFinishMovementService
    {
        private readonly IFinishMovementDataAccess _dataAccess;

        public FinishMovementService(IFinishMovementDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ItemViewModel>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();

        public Task<List<FinishFromLocationModel>> GetFromLocationsAsync(string itemId) => _dataAccess.GetFromLocationsAsync(itemId);

        public Task<List<SFTargetStoreModel>> GetStoresAsync(string userName) => _dataAccess.GetStoresAsync(userName);

        public Task<List<SFTargetRackModel>> GetRacksAsync(int storeRefId) => _dataAccess.GetRacksAsync(storeRefId);

        public Task<List<SFTargetShelfModel>> GetShelvesAsync(int storeRefId, string rackNo) => _dataAccess.GetShelvesAsync(storeRefId, rackNo);

        public Task<double> GetTargetShelfQtyAsync(string itemId, int shelfRefId) => _dataAccess.GetTargetShelfQtyAsync(itemId, shelfRefId);

        public Task<bool> ExecuteMovementAsync(FinishMovementRequest request) => _dataAccess.ExecuteMovementAsync(request);
    }
}
