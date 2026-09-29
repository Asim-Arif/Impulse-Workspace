using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IOrderLotsTrackingDataAccess
    {
        Task<List<TrackingCustomerLookupItem>> GetCustomersLookupAsync();
        Task<List<TrackingOrderLookupItem>> GetOrdersLookupAsync(string? custCode = null);
        Task<List<string>> GetHubNamesLookupAsync();
        Task<OrderLotsTrackingDashboardDto> GetOrderLotsTrackingAsync(OrderLotsTrackingFilter filter);
    }
}
