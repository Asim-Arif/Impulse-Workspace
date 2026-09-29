using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IOrderLotsTrackingService
    {
        Task<List<TrackingCustomerLookupItem>> GetCustomersLookupAsync();
        Task<List<TrackingOrderLookupItem>> GetOrdersLookupAsync(string? custCode = null);
        Task<List<string>> GetHubNamesLookupAsync();
        Task<OrderLotsTrackingDashboardDto> GetOrderLotsTrackingAsync(OrderLotsTrackingFilter filter);
    }
}
