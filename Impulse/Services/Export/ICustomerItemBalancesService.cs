using DataAccessLibrary.Models.ViewModels.Export;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Export
{
    public interface ICustomerItemBalancesService
    {
        Task<List<CustomerItemBalancesViewModel>> GetCustomerItemBalancesAsync(string custCode);
        Task UpdateBalanceQuantityAsync(long id, decimal qtyDiff);
        Task ToggleDeliveryStatusAsync(long id, int newStatus);
    }
}
