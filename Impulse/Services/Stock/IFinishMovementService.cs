using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IFinishMovementService
    {
        Task<List<ItemViewModel>> GetArticlesAsync();
        Task<List<FinishFromLocationModel>> GetFromLocationsAsync(string itemId);
        Task<List<SFTargetStoreModel>> GetStoresAsync(string userName);
        Task<List<SFTargetRackModel>> GetRacksAsync(int storeRefId);
        Task<List<SFTargetShelfModel>> GetShelvesAsync(int storeRefId, string rackNo);
        Task<double> GetTargetShelfQtyAsync(string itemId, int shelfRefId);
        Task<bool> ExecuteMovementAsync(FinishMovementRequest request);
    }
}
