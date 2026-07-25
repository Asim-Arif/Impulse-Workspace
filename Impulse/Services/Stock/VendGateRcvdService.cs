using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Services.Stock
{
    public class VendGateRcvdService : IVendGateRcvdService
    {
        private readonly IVendGateRcvdDataAccess _dataAccess;

        public VendGateRcvdService(IVendGateRcvdDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<object>> GetVendorsWithPendingOrdersAsync()
        {
            return await _dataAccess.GetVendorsWithPendingOrdersAsync();
        }

        public async Task<List<string>> GetPendingOrdersAsync(string vendId)
        {
            return await _dataAccess.GetPendingOrdersAsync(vendId);
        }

        public async Task<List<VendGateRcvdItemViewModel>> GetPendingOrderItemsAsync(string orderNo)
        {
            return await _dataAccess.GetPendingOrderItemsAsync(orderNo);
        }

        public async Task<List<string>> GetStoresAsync()
        {
            return await _dataAccess.GetStoresAsync();
        }

        public async Task<string> SaveVendorReceivingAsync(VendGateRcvdSaveRequest request, string currentUser)
        {
            return await _dataAccess.SaveVendorReceivingAsync(request, currentUser);
        }
    }
}
