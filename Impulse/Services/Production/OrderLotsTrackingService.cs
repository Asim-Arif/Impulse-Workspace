using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Production
{
    public class OrderLotsTrackingService : IOrderLotsTrackingService
    {
        private readonly IOrderLotsTrackingDataAccess _dataAccess;
        private readonly ILogger<OrderLotsTrackingService> _logger;

        public OrderLotsTrackingService(IOrderLotsTrackingDataAccess dataAccess, ILogger<OrderLotsTrackingService> logger)
        {
            _dataAccess = dataAccess;
            _logger = logger;
        }

        public async Task<List<TrackingCustomerLookupItem>> GetCustomersLookupAsync()
        {
            try
            {
                return await _dataAccess.GetCustomersLookupAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching customers lookup for Order Lots Tracking");
                return new List<TrackingCustomerLookupItem>();
            }
        }

        public async Task<List<TrackingOrderLookupItem>> GetOrdersLookupAsync(string? custCode = null)
        {
            try
            {
                return await _dataAccess.GetOrdersLookupAsync(custCode);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching orders lookup for Order Lots Tracking");
                return new List<TrackingOrderLookupItem>();
            }
        }

        public async Task<List<string>> GetHubNamesLookupAsync()
        {
            try
            {
                return await _dataAccess.GetHubNamesLookupAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching hub names lookup for Order Lots Tracking");
                return new List<string>();
            }
        }

        public async Task<OrderLotsTrackingDashboardDto> GetOrderLotsTrackingAsync(OrderLotsTrackingFilter filter)
        {
            try
            {
                return await _dataAccess.GetOrderLotsTrackingAsync(filter);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching order lots tracking data");
                throw;
            }
        }
    }
}
