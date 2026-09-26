using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IPpcStockAdjustmentDataAccess
    {
        Task<PpcStockAdjustmentOrderHeaderModel?> GetOrderHeaderAsync(string orderNo);
        Task<List<PpcStockAdjustmentOrderHeaderModel>> GetActiveOrdersWithStockPlanningAsync();
        Task<List<PpcStockAdjustmentPlannedItemModel>> GetPlannedStockItemsAsync(string orderNo);
        Task<List<PpcStockShelfAllocationModel>> GetAvailableShelfLotsAsync(string itemId);
        Task<List<ItemProcessLookupModel>> GetProcessesForItemAsync(string itemId);
        Task<PpcStockAdjustmentResultModel> SaveStockAdjustmentAsync(PpcStockAdjustmentSubmissionModel submission);
    }
}
