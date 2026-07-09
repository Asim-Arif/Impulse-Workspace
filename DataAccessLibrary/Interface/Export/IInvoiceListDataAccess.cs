using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface IInvoiceListDataAccess
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
    }
}
