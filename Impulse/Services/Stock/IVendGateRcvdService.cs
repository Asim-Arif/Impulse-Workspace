using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Services.Stock
{
    public interface IVendGateRcvdService
    {
        Task<List<object>> GetVendorsWithPendingOrdersAsync();
        Task<List<string>> GetPendingOrdersAsync(string vendId);
        Task<List<VendGateRcvdItemViewModel>> GetPendingOrderItemsAsync(string orderNo);
        Task<List<string>> GetStoresAsync();
        Task<string> SaveVendorReceivingAsync(VendGateRcvdSaveRequest request, string currentUser);
    }
}
