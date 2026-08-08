using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Dashboard;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Dashboard
{
    public interface IProductionPlanningDashboardService
    {
        Task<List<LookupItemString>> GetCustomersAsync();
        Task<List<string>> GetCountriesByCustomerAsync(string custCode);
        Task<List<LookupItemString>> GetArticlesAsync(string custCode, string country);
        Task<ProductionPlanningDashboardData> GetDashboardDataAsync(string itemCode, string? filterOrderNo = null);
        Task<List<ForgingCrossDemandItem>> GetForgingCrossDemandAsync(long rmid, string currentItemCode);
        Task<bool> ToggleForgingFunctionalStatusAsync(string itemCode, long rmid, bool isInactive);
        Task<bool> CloseLotAsync(string lotNo, string userName, string machineName);
        Task<bool> ClosePOAsync(long entryID);
    }
}
