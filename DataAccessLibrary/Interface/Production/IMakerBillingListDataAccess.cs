using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IMakerBillingListDataAccess
    {
        Task<List<LookupItemInt>> GetMakersAsync();
        Task<List<MakerBillingListItem>> GetPostedBillsListAsync(MakerBillingListFilter filter);
        Task<List<LookupItemString>> GetPaymentCreditAccountsAsync();
        Task<string?> GetMakerAccountNoAsync(int vendId);
        Task<bool> RecordPaymentVoucherAsync(long mpbRefId, string voucherNo, decimal paidAmt);
        Task<bool> DeletePostedBillAsync(long entryId, string userName);
    }
}
