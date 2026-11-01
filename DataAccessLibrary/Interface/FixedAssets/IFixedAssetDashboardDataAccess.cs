using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.FixedAssets;

namespace DataAccessLibrary.Interface.FixedAssets
{
    public interface IFixedAssetDashboardDataAccess
    {
        Task<FixedAssetDashboardResult> GetDashboardDataAsync(FixedAssetDashboardFilter filter, CancellationToken cancellationToken = default);
    }
}

