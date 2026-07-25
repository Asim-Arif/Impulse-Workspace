using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class MaterialPlacementService : IMaterialPlacementService
    {
        private readonly IMaterialPlacementDataAccess _dataAccess;

        public MaterialPlacementService(IMaterialPlacementDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<MaterialPlacementListViewModel>> GetPendingPlacementsAsync(string vendorId, string materialId)
        {
            return _dataAccess.GetPendingPlacementsAsync(vendorId, materialId);
        }

        public Task<List<StoreViewModel>> GetStoresForUserAsync(int userId)
        {
            return _dataAccess.GetStoresForUserAsync(userId);
        }

        public Task<List<RackViewModel>> GetRacksAsync(int storeId)
        {
            return _dataAccess.GetRacksAsync(storeId);
        }

        public Task<List<ShelfViewModel>> GetShelvesAsync(int rackId)
        {
            return _dataAccess.GetShelvesAsync(rackId);
        }

        public Task SavePlacementAsync(PlacementSaveRequest request, string userName, string machineName)
        {
            return _dataAccess.SavePlacementAsync(request, userName, machineName);
        }
    }
}
