using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IGateDashboardDataAccess
{
    Task<GateDashboardModel> GetDashboardAsync(CancellationToken cancellationToken = default);
}

