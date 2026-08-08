using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Dashboard;

namespace DataAccessLibrary.Interface.Dashboard
{
    public interface ICommandCenterDashboardDataAccess
    {
        Task<CommandCenterDashboardData> GetCommandCenterDataAsync();
        Task<List<StockGroupSummaryItem>> GetStockGroupSummaryAsync(int stockType);
        Task<List<StockGroupItemDetail>> GetStockGroupDetailAsync(int stockType, int groupID);
    }
}
