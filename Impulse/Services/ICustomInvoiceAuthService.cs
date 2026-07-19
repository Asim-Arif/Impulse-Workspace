using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services
{
    public interface ICustomInvoiceAuthService
    {
        Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync();
        Task<IEnumerable<string>> GetCountriesAsync(string? customerCode);
        Task<IEnumerable<ChartOfAccountsModel>> GetSalesAccountsAsync();
        Task<IEnumerable<CustomInvoiceAuthItemModel>> GetPendingInvoicesAsync(string? customerCode, string? country, DateTime fromDate, DateTime toDate);
        Task<string> GetExchDiffAccNoAsync();
        Task<int> AuthorizeInvoicesAsync(DateTime voucherDate, string salesAccNo, string exchDiffAccNo, List<CustomInvoiceAuthItemModel> items);
    }
}
