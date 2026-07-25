using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class VendorBillingListService : IVendorBillingListService
    {
        private readonly IVendorBillingListDataAccess _dataAccess;

        public VendorBillingListService(IVendorBillingListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<IEnumerable<VendorBillingListViewModel>> GetVendorBillingListAsync(DateTime fromDate, DateTime toDate, string vendorId, bool isSample)
        {
            return _dataAccess.GetVendorBillingListAsync(fromDate, toDate, vendorId, isSample);
        }

        public Task<IEnumerable<object>> GetVendorsAsync()
        {
            return _dataAccess.GetVendorsAsync();
        }

        public Task DeleteVendorBillAsync(int entryId)
        {
            return _dataAccess.DeleteVendorBillAsync(entryId);
        }
    }
}
