using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class VendOrderService : IVendOrderService
    {
        private readonly IVendOrderDataAccess _vendOrderDataAccess;

        public VendOrderService(IVendOrderDataAccess vendOrderDataAccess)
        {
            _vendOrderDataAccess = vendOrderDataAccess;
        }

        public Task<VendOrderViewModel> GetVendOrderAsync(string orderNo)
        {
            return _vendOrderDataAccess.GetVendOrderAsync(orderNo);
        }

        public async Task<string> SaveVendOrderAsync(VendOrderViewModel model)
        {
            if (string.IsNullOrWhiteSpace(model.OrderNo))
            {
                return await _vendOrderDataAccess.InsertVendOrderAsync(model);
            }
            else
            {
                // Check if it exists? We assume it exists if OrderNo is provided.
                await _vendOrderDataAccess.UpdateVendOrderAsync(model);
                return model.OrderNo;
            }
        }

        public Task<bool> DeleteVendOrderAsync(string orderNo)
        {
            return _vendOrderDataAccess.DeleteVendOrderAsync(orderNo);
        }

        public Task<string> GetNextOrderNoAsync()
        {
            return _vendOrderDataAccess.GetNextOrderNoAsync();
        }
    }
}
