using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface IInvoiceListService
    {
        Task<IEnumerable<CompanyLookupModel>> GetCompaniesAsync();
        Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync();
        Task<IEnumerable<string>> GetCountriesAsync();
        Task<IEnumerable<InvoiceListModel>> GetInvoicesAsync(
            int invoiceType,
            long companyId,
            string customerCode,
            string country,
            string dateRange,
            DateTime? fromDate,
            DateTime? toDate,
            string searchText);
        Task DeleteInvoiceAsync(string customInvoiceNo);
    }
}
