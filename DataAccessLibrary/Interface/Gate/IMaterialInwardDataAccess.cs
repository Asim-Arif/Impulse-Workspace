using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IMaterialInwardDataAccess
{
    Task<IReadOnlyList<MaterialInwardListRowModel>> GetEntriesAsync(MaterialInwardFilterModel filter, CancellationToken cancellationToken = default);
    Task<MaterialInwardEntryModel?> GetEntryAsync(int entryId, CancellationToken cancellationToken = default);
    Task<int> SaveAsync(MaterialInwardEntryModel model, string userName, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<VehicleEntryListRowModel>> GetVehicleOptionsAsync(CancellationToken cancellationToken = default);
    Task DeleteAsync(int entryId, CancellationToken cancellationToken = default);
}

