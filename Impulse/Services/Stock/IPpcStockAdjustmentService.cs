using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Services.Stock
{
    public interface IPpcStockAdjustmentService
    {
        Task<PpcStockAdjustmentOrderHeaderModel?> GetOrderHeaderAsync(string orderNo);
        Task<List<PpcStockAdjustmentOrderHeaderModel>> GetActiveOrdersWithStockPlanningAsync();
        Task<List<PpcStockAdjustmentPlannedItemModel>> GetPlannedStockItemsAsync(string orderNo);
        Task AutoAllocateStockAsync(PpcStockAdjustmentPlannedItemModel item);
        Task<PpcStockAdjustmentResultModel> SaveStockAdjustmentAsync(PpcStockAdjustmentSubmissionModel submission);
    }
}
