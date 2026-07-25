using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IRMMovementDataAccess
    {
        Task<List<LocationwiseStatusViewModel>> GetLocationsForMaterialAsync(string materialId);
        Task<double> GetShelfQuantityAsync(int shelfId, string materialId);
        Task SaveMovementAsync(RMMovementSaveRequest request, string userName, string machineName);
    }
}
