using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Accounts
{
    public interface ICustomInvoiceAuthDataAccess
    {
        Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync();
        Task<IEnumerable<string>> GetCountriesAsync(string? customerCode);
        Task<IEnumerable<ChartOfAccountsModel>> GetSalesAccountsAsync();
        Task<IEnumerable<CustomInvoiceAuthItemModel>> GetPendingInvoicesAsync(string? customerCode, string? country, DateTime fromDate, DateTime toDate);
        Task<int> AuthorizeInvoicesAsync(CustomInvoiceAuthSaveModel request);
    }
}
