using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class OrderManagementService : IOrderManagementService
    {
        private readonly IOrderManagementDataAccess _dataAccess;
        private readonly ILogger<OrderManagementService> _logger;

        public OrderManagementService(IOrderManagementDataAccess dataAccess, ILogger<OrderManagementService> logger)
        {
            _dataAccess = dataAccess;
            _logger = logger;
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
    }
}
