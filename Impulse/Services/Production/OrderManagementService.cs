using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using Microsoft.AspNetCore.Hosting;
using System.IO;

namespace Impulse.Services.Production
{
    public class OrderManagementService : IOrderManagementService
    {
        private readonly IOrderManagementDataAccess _dataAccess;
        private readonly ILogger<OrderManagementService> _logger;
        private readonly IWebHostEnvironment _environment;

        public OrderManagementService(
            IOrderManagementDataAccess dataAccess, 
            ILogger<OrderManagementService> logger,
            IWebHostEnvironment environment)
        {
            _dataAccess = dataAccess;
            _logger = logger;
            _environment = environment;
        }

        public async Task<List<LookupItemString>> GetCustomersAsync()
        {
            try
            {
                return await _dataAccess.GetCustomersAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching customers for Order Management");
                return new List<LookupItemString>();
            }
        }

        public async Task<List<CustomerOrderHeaderDto>> GetOrdersAsync(OrderManagementFilter filter)
        {
            try
            {
                return await _dataAccess.GetOrdersAsync(filter);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching orders for Order Management");
                return new List<CustomerOrderHeaderDto>();
            }
        }

        public OrderSummaryCardDto CalculateOrderSummaryMetrics(List<CustomerOrderHeaderDto> orders)
        {
            try
            {
                return _dataAccess.CalculateOrderSummaryMetrics(orders);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error calculating order summary metrics from orders");
                return new OrderSummaryCardDto();
            }
        }

        public async Task<OrderSummaryCardDto> GetOrderSummaryMetricsAsync(OrderManagementFilter filter)
        {
            try
            {
                return await _dataAccess.GetOrderSummaryMetricsAsync(filter);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error calculating order summary metrics");
                return new OrderSummaryCardDto();
            }
        }

        public async Task<List<OrderItemProgressDto>> GetOrderItemsAsync(string orderNo)
        {
            try
            {
                return await _dataAccess.GetOrderItemsAsync(orderNo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching order items for Order {OrderNo}", orderNo);
                return new List<OrderItemProgressDto>();
            }
        }

        public async Task<List<ItemPurchaseOrderDto>> GetItemPurchaseOrdersAsync(string orderNo, string itemCode, string? compItemCode = null)
        {
            try
            {
                return await _dataAccess.GetItemPurchaseOrdersAsync(orderNo, itemCode, compItemCode);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching POs for Order {OrderNo}, Item {ItemCode}", orderNo, itemCode);
                return new List<ItemPurchaseOrderDto>();
            }
        }

        public async Task<List<ItemRunningLotDto>> GetItemRunningLotsAsync(string orderNo, string itemCode, string? compItemCode = null)
        {
            try
            {
                return await _dataAccess.GetItemRunningLotsAsync(orderNo, itemCode, compItemCode);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching running lots for Order {OrderNo}, Item {ItemCode}", orderNo, itemCode);
                return new List<ItemRunningLotDto>();
            }
        }

        public async Task<List<ItemStockAdjustmentDto>> GetItemStockAdjustmentsAsync(string orderNo, string itemCode, string? compItemCode = null)
        {
            try
            {
                return await _dataAccess.GetItemStockAdjustmentsAsync(orderNo, itemCode, compItemCode);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching stock adjustments for Order {OrderNo}, Item {ItemCode}", orderNo, itemCode);
                return new List<ItemStockAdjustmentDto>();
            }
        }

        public async Task<List<ItemDispatchDetailDto>> GetItemDispatchDetailsAsync(string orderNo, string itemCode, string? compItemCode = null)
        {
            try
            {
                return await _dataAccess.GetItemDispatchDetailsAsync(orderNo, itemCode, compItemCode);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching dispatch details for Order {OrderNo}, Item {ItemCode}", orderNo, itemCode);
                return new List<ItemDispatchDetailDto>();
            }
        }

        public async Task<(bool Success, string? ErrorMessage, string? RelativePath)> UploadMasterPoSignedCopyAsync(
            string masterPoNo, int entryId, Stream fileStream, string originalFileName, string userName)
        {
            try
            {
                if (fileStream == null)
                    return (false, "File stream is empty.", null);

                var ext = Path.GetExtension(originalFileName);
                if (string.IsNullOrWhiteSpace(ext) || !ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                    return (false, "Only PDF (.pdf) files are allowed for Maker Signed Copy.", null);

                var webRoot = _environment.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
                var folderPath = Path.Combine(webRoot, "uploads", "maker_po_signed");
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                // Sanitize Master PO number for safe filename
                var cleanMasterPo = !string.IsNullOrWhiteSpace(masterPoNo) 
                    ? string.Concat(masterPoNo.Split(Path.GetInvalidFileNameChars())) 
                    : $"Entry_{entryId}";

                var uniqueFileName = $"Signed_MasterPO_{cleanMasterPo}_{DateTime.Now:yyyyMMddHHmmss}_{Guid.NewGuid().ToString().Substring(0, 8)}.pdf";
                var physicalFilePath = Path.Combine(folderPath, uniqueFileName);

                using (var destStream = new FileStream(physicalFilePath, FileMode.Create))
                {
                    await fileStream.CopyToAsync(destStream);
                }

                var relativePath = $"/uploads/maker_po_signed/{uniqueFileName}";

                var dbSuccess = await _dataAccess.SaveMasterPoSignedCopyAsync(
                    masterPoNo, entryId, relativePath, originalFileName, userName);

                if (!dbSuccess)
                {
                    // Rollback physical file if DB fails
                    if (File.Exists(physicalFilePath))
                    {
                        try { File.Delete(physicalFilePath); } catch { }
                    }
                    return (false, "Failed to update database record with signed copy path.", null);
                }

                return (true, null, relativePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading signed copy for Master PO {MasterPONo} / EntryID {EntryID}", masterPoNo, entryId);
                return (false, $"Upload failed: {ex.Message}", null);
            }
        }

        public async Task<bool> DeleteMasterPoSignedCopyAsync(string masterPoNo, int entryId, string? existingFilePath)
        {
            try
            {
                var dbSuccess = await _dataAccess.DeleteMasterPoSignedCopyAsync(masterPoNo, entryId);
                if (dbSuccess && !string.IsNullOrWhiteSpace(existingFilePath))
                {
                    var webRoot = _environment.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
                    var cleanPath = existingFilePath.TrimStart('/').Replace('/', Path.DirectorySeparatorChar);
                    var physicalPath = Path.Combine(webRoot, cleanPath);
                    if (File.Exists(physicalPath))
                    {
                        try { File.Delete(physicalPath); } catch { }
                    }
                }
                return dbSuccess;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting signed copy for Master PO {MasterPONo} / EntryID {EntryID}", masterPoNo, entryId);
                return false;
            }
        }
    }
}
