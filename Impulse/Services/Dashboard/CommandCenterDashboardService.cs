using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Dashboard;
using DataAccessLibrary.Models.ViewModels.Dashboard;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Dashboard
{
    public class CommandCenterDashboardService : ICommandCenterDashboardService
    {
        private readonly ICommandCenterDashboardDataAccess _dataAccess;
        private readonly ILogger<CommandCenterDashboardService> _logger;

        public CommandCenterDashboardService(
            ICommandCenterDashboardDataAccess dataAccess,
            ILogger<CommandCenterDashboardService> logger)
        {
            _dataAccess = dataAccess;
            _logger = logger;
        }

        public async Task<CommandCenterDashboardData> GetCommandCenterDataAsync()
        {
            try
            {
                return await _dataAccess.GetCommandCenterDataAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Command Center Dashboard data.");
                return new CommandCenterDashboardData();
            }
        }

        public async Task<List<StockGroupSummaryItem>> GetStockGroupSummaryAsync(int stockType)
        {
            try
            {
                return await _dataAccess.GetStockGroupSummaryAsync(stockType);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Stock Group Summary for Type {StockType}", stockType);
                return new List<StockGroupSummaryItem>();
            }
        }

        public async Task<List<StockGroupItemDetail>> GetStockGroupDetailAsync(int stockType, int groupID)
        {
            try
            {
                return await _dataAccess.GetStockGroupDetailAsync(stockType, groupID);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Stock Group Details for Type {StockType}, GroupID {GroupID}", stockType, groupID);
                return new List<StockGroupItemDetail>();
            }
        }
    }
}
