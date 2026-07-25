using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IRMMovementService
    {
        Task<List<LocationwiseStatusViewModel>> GetLocationsForMaterialAsync(string materialId);
        Task<double> GetShelfQuantityAsync(int shelfId, string materialId);
        Task SaveMovementAsync(RMMovementSaveRequest request, string userName, string machineName);
    }
}
