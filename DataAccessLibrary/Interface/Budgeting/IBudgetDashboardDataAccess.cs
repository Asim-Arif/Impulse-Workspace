using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Budgeting;

namespace DataAccessLibrary.Interface.Budgeting
{
    public interface IBudgetDashboardDataAccess
    {
        Task<BudgetDashboardResult> GetDashboardDataAsync(BudgetDashboardFilter filter, CancellationToken cancellationToken = default);
    }
}

