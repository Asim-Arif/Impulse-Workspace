using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface ISFMovementDataAccess
    {
        Task<List<ItemViewModel>> GetArticlesAsync();
        Task<List<ItemProcessViewModel>> GetProcessesByArticleAsync(string itemId);
        Task<List<SFTargetStoreModel>> GetFromStoresAsync(string itemId, int processId);
        Task<List<SFFromLocationModel>> GetFromLocationsAsync(string itemId, int processId, int storeRefId);
        Task<List<SFTargetStoreModel>> GetTargetStoresAsync(string userName);
        Task<List<SFTargetRackModel>> GetTargetRacksAsync(int storeRefId);
        Task<List<SFTargetShelfModel>> GetTargetShelvesAsync(int storeRefId, string rackNo);
        Task<int> GetTargetLocationQtyAsync(string itemId, int processId, int shelfRefId);
        Task<bool> TransferLocationAsync(SFMovementRequest request);
    }
}
