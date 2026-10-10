using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Services.Stock
{
    public class VendGateRcvdService : IVendGateRcvdService
    {
        private readonly IVendGateRcvdDataAccess _dataAccess;
        private readonly IStockNotificationService _stockNotificationService;

        public VendGateRcvdService(
            IVendGateRcvdDataAccess dataAccess,
            IStockNotificationService stockNotificationService)
        {
            _dataAccess = dataAccess;
            _stockNotificationService = stockNotificationService;
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
            var rcvNo = await _dataAccess.SaveVendorReceivingAsync(request, currentUser);

            if (!string.IsNullOrWhiteSpace(rcvNo) && request?.ReceivedItems != null)
            {
                var processedCodes = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                foreach (var item in request.ReceivedItems)
                {
                    if (!string.IsNullOrWhiteSpace(item.MaterialID) && processedCodes.Add(item.MaterialID.Trim()))
                    {
                        _ = _stockNotificationService.CheckAndNotifyStockMovementAsync(
                            item.MaterialID.Trim(),
                            "Vendor Receiving",
                            rcvNo);
                    }
                }
            }

            return rcvNo;
        }
    }
}
