using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class VendorBillingService : IVendorBillingService
    {
        private readonly IVendorBillingDataAccess _dataAccess;

        public VendorBillingService(IVendorBillingDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<IEnumerable<VendorLookupModel>> GetVendorsAsync() =>
            _dataAccess.GetVendorsAsync();

        public Task<IEnumerable<AccountLookupModel>> GetActiveAccountsAsync() =>
            _dataAccess.GetActiveAccountsAsync();

        public Task<IEnumerable<VendorBillingRowViewModel>> GetUnpostedReceivingsAsync(
            string vendorAccNo, DateTime dateFrom, DateTime dateTo) =>
            _dataAccess.GetUnpostedReceivingsAsync(vendorAccNo, dateFrom, dateTo);

        public Task<IEnumerable<MaterialDeductionRowViewModel>> GetMaterialIssuancesAsync(
            string vendorAccNo) =>
            _dataAccess.GetMaterialIssuancesAsync(vendorAccNo);

        public Task<decimal> GetShortTermDeductionAsync(string vendorAccNo) =>
            _dataAccess.GetShortTermDeductionAsync(vendorAccNo);

        public Task<decimal> GetLongTermDeductionAsync(string vendorAccNo) =>
            _dataAccess.GetLongTermDeductionAsync(vendorAccNo);

        public Task<int> GetNextBillNoAsync(string vendorAccNo) =>
            _dataAccess.GetNextBillNoAsync(vendorAccNo);

        public Task UpdateReceivingRateAsync(int entryId, float newRate) =>
            _dataAccess.UpdateReceivingRateAsync(entryId, newRate);

        public Task HideFromBillingAsync(int entryId, string userName, string machineName) =>
            _dataAccess.HideFromBillingAsync(entryId, userName, machineName);

        public Task SavePostedBillAsync(VendorBillingPostModel model) =>
            _dataAccess.SavePostedBillAsync(model);
    }
}
