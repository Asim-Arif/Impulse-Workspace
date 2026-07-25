using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IMaterialPlacementDataAccess
    {
        Task<List<MaterialPlacementListViewModel>> GetPendingPlacementsAsync(string vendorId, string materialId);
        Task<List<StoreViewModel>> GetStoresForUserAsync(int userId);
        Task<List<RackViewModel>> GetRacksAsync(int storeId);
        Task<List<ShelfViewModel>> GetShelvesAsync(int rackId);
        Task SavePlacementAsync(PlacementSaveRequest request, string userName, string machineName);
    }
}
