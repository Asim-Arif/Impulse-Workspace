using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Export
{
    public class CustomerItemBalancesService : ICustomerItemBalancesService
    {
        private readonly ICustomerItemBalancesDataAccess _dataAccess;

        public CustomerItemBalancesService(ICustomerItemBalancesDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<CustomerItemBalancesViewModel>> GetCustomerItemBalancesAsync(string custCode)
        {
            return await _dataAccess.GetCustomerItemBalancesAsync(custCode);
        }

        public async Task UpdateBalanceQuantityAsync(long id, decimal qtyDiff)
        {
            await _dataAccess.UpdateBalanceQuantityAsync(id, qtyDiff);
        }

        public async Task ToggleDeliveryStatusAsync(long id, int newStatus)
        {
            await _dataAccess.ToggleDeliveryStatusAsync(id, newStatus);
        }
    }
}
