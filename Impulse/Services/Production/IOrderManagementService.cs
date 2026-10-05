using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IOrderManagementService
    {
        Task<List<LookupItemString>> GetCustomersAsync();
        Task<List<CustomerOrderHeaderDto>> GetOrdersAsync(OrderManagementFilter filter);
        OrderSummaryCardDto CalculateOrderSummaryMetrics(List<CustomerOrderHeaderDto> orders);
        Task<OrderSummaryCardDto> GetOrderSummaryMetricsAsync(OrderManagementFilter filter);
        Task<List<OrderItemProgressDto>> GetOrderItemsAsync(string orderNo);
        Task<List<ItemPurchaseOrderDto>> GetItemPurchaseOrdersAsync(string orderNo, string itemCode, string? compItemCode = null);
        Task<List<ItemRunningLotDto>> GetItemRunningLotsAsync(string orderNo, string itemCode, string? compItemCode = null);
        Task<List<ItemStockAdjustmentDto>> GetItemStockAdjustmentsAsync(string orderNo, string itemCode, string? compItemCode = null);
        Task<List<ItemDispatchDetailDto>> GetItemDispatchDetailsAsync(string orderNo, string itemCode, string? compItemCode = null);
        Task<(bool Success, string? ErrorMessage, string? RelativePath)> UploadMasterPoSignedCopyAsync(string masterPoNo, int entryId, System.IO.Stream fileStream, string originalFileName, string userName);
        Task<bool> DeleteMasterPoSignedCopyAsync(string masterPoNo, int entryId, string? existingFilePath);
    }
}
