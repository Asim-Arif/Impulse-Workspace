using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Dashboard;
using DataAccessLibrary.Models.ViewModels.Dashboard;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Dashboard
{
    public class ProductionPlanningDashboardService : IProductionPlanningDashboardService
    {
        private readonly IProductionPlanningDashboardDataAccess _dac;

        public ProductionPlanningDashboardService(IProductionPlanningDashboardDataAccess dac)
        {
            _dac = dac;
        }

        public Task<List<LookupItemString>> GetCustomersAsync() => _dac.GetCustomersAsync();

        public Task<List<string>> GetCountriesByCustomerAsync(string custCode) => _dac.GetCountriesByCustomerAsync(custCode);

        public Task<List<LookupItemString>> GetArticlesAsync(string custCode, string country) => _dac.GetArticlesAsync(custCode, country);

        public Task<ProductionPlanningDashboardData> GetDashboardDataAsync(string itemCode, string? filterOrderNo = null) =>
            _dac.GetDashboardDataAsync(itemCode, filterOrderNo);

        public Task<List<ForgingCrossDemandItem>> GetForgingCrossDemandAsync(long rmid, string currentItemCode) =>
            _dac.GetForgingCrossDemandAsync(rmid, currentItemCode);

        public Task<bool> ToggleForgingFunctionalStatusAsync(string itemCode, long rmid, bool isInactive) =>
            _dac.ToggleForgingFunctionalStatusAsync(itemCode, rmid, isInactive);

        public Task<bool> CloseLotAsync(string lotNo, string userName, string machineName) =>
            _dac.CloseLotAsync(lotNo, userName, machineName);

        public Task<bool> ClosePOAsync(long entryID) => _dac.ClosePOAsync(entryID);
    }
}
