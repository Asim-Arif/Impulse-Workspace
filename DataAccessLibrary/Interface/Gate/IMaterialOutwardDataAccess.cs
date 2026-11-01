using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IMaterialOutwardDataAccess
{
    Task<IReadOnlyList<MaterialOutwardListRowModel>> GetEntriesAsync(MaterialOutwardFilterModel filter, CancellationToken cancellationToken = default);
    Task<MaterialOutwardEntryModel?> GetEntryAsync(int entryId, CancellationToken cancellationToken = default);
    Task<int> SaveAsync(MaterialOutwardEntryModel model, string userName, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<VehicleEntryListRowModel>> GetVehicleOptionsAsync(CancellationToken cancellationToken = default);
    Task DeleteAsync(int entryId, CancellationToken cancellationToken = default);
}

