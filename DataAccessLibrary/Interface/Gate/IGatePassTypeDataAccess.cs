using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IGatePassTypeDataAccess
{
    Task<IReadOnlyList<GatePassTypeModel>> GetAsync(CancellationToken cancellationToken = default);
    Task<int> SaveAsync(GatePassTypeModel model, CancellationToken cancellationToken = default);
    Task SetActiveAsync(int entryId, bool active, CancellationToken cancellationToken = default);
    Task DeleteAsync(int entryId, CancellationToken cancellationToken = default);
}

