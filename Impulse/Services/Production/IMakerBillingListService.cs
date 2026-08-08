using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IMakerBillingListService
    {
        Task<List<LookupItemInt>> GetMakersAsync();
        Task<List<MakerBillingListItem>> GetPostedBillsListAsync(MakerBillingListFilter filter);
        Task<List<LookupItemString>> GetPaymentCreditAccountsAsync();
        Task<string?> GetMakerAccountNoAsync(int vendId);
        Task<string> PostPaymentVoucherAsync(MakerBillingPaymentVoucherRequest request);
        Task<bool> DeletePostedBillAsync(long entryId, string userName);
    }
}
