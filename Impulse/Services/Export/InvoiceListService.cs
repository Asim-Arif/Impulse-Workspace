using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;

namespace Impulse.Services.Export
{
    public class InvoiceListService : IInvoiceListService
    {
        private readonly IInvoiceListDataAccess _dataAccess;
        private readonly IDBHelperService _dbHelper;

        public InvoiceListService(IInvoiceListDataAccess dataAccess, IDBHelperService dbHelper)
        {
            _dataAccess = dataAccess;
            _dbHelper = dbHelper;
        }

        public Task<IEnumerable<CompanyLookupModel>> GetCompaniesAsync() => _dataAccess.GetCompaniesAsync();
        public Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        public Task<IEnumerable<string>> GetCountriesAsync() => _dataAccess.GetCountriesAsync();

        public Task<IEnumerable<InvoiceListModel>> GetInvoicesAsync(
            int invoiceType,
            long companyId,
            string customerCode,
            string country,
            string dateRange,
            DateTime? fromDate,
            DateTime? toDate,
            string searchText)
        {
            return _dataAccess.GetInvoicesAsync(invoiceType, companyId, customerCode, country, dateRange, fromDate, toDate, searchText);
        }

        public async Task DeleteInvoiceAsync(string customInvoiceNo)
        {
            await _dbHelper.ExecuteAsync("DELETE FROM dbo.CustomInvoice WHERE CustomInvoice = @Invoice", new { Invoice = customInvoiceNo });
            // Should also delete associated details in a real scenario, but matching legacy behavior for now.
        }
    }
}
