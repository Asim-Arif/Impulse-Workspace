using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class RMMovementService : IRMMovementService
    {
        private readonly IRMMovementDataAccess _dataAccess;

        public RMMovementService(IRMMovementDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<LocationwiseStatusViewModel>> GetLocationsForMaterialAsync(string materialId)
        {
            return _dataAccess.GetLocationsForMaterialAsync(materialId);
        }

        public Task<double> GetShelfQuantityAsync(int shelfId, string materialId)
        {
            return _dataAccess.GetShelfQuantityAsync(shelfId, materialId);
        }

        public Task SaveMovementAsync(RMMovementSaveRequest request, string userName, string machineName)
        {
            return _dataAccess.SaveMovementAsync(request, userName, machineName);
        }
    }
}
