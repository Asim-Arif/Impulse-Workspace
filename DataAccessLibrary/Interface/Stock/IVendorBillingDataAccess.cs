using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IVendorBillingDataAccess
    {
        Task<IEnumerable<VendorLookupModel>> GetVendorsAsync();
        Task<IEnumerable<AccountLookupModel>> GetActiveAccountsAsync();
        Task<IEnumerable<VendorBillingRowViewModel>> GetUnpostedReceivingsAsync(string vendorAccNo, System.DateTime dateFrom, System.DateTime dateTo);
        Task<IEnumerable<MaterialDeductionRowViewModel>> GetMaterialIssuancesAsync(string vendorAccNo);
        Task<decimal> GetShortTermDeductionAsync(string vendorAccNo);
        Task<decimal> GetLongTermDeductionAsync(string vendorAccNo);
        Task<int> GetNextBillNoAsync(string vendorAccNo);
        Task UpdateReceivingRateAsync(int entryId, float newRate);
        Task HideFromBillingAsync(int entryId, string userName, string machineName);
        Task SavePostedBillAsync(VendorBillingPostModel model);
    }
}
