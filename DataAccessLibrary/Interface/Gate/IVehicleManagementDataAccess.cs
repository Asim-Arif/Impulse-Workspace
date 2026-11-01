using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IVehicleManagementDataAccess
{
    Task<IReadOnlyList<VehicleCategoryModel>> GetCategoriesAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<VehicleEntryListRowModel>> GetVehiclesAsync(VehicleFilterModel filter, CancellationToken cancellationToken = default);
    Task<VehicleEntryModel?> GetVehicleAsync(int entryId, CancellationToken cancellationToken = default);
    Task<int> SaveVehicleAsync(VehicleEntryModel model, string userName, CancellationToken cancellationToken = default);
    Task DeleteVehicleAsync(int entryId, CancellationToken cancellationToken = default);
}

