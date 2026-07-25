using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IVendorBillingListDataAccess
    {
        Task<IEnumerable<VendorBillingListViewModel>> GetVendorBillingListAsync(DateTime fromDate, DateTime toDate, string vendorId, bool isSample);
        Task<IEnumerable<object>> GetVendorsAsync();
        Task DeleteVendorBillAsync(int entryId);
    }
}
