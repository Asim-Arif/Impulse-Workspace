using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Export
{
    public interface IProformaListService
    {
        Task<IEnumerable<ProformaListModel>> GetProformaListAsync(DateTime fromDate, DateTime toDate, string custCode, long companyId, string country);
        Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync();
        Task<IEnumerable<CompanyLookupModel>> GetCompaniesAsync();
        Task<IEnumerable<string>> GetCountriesAsync(string custCode);
        Task<bool> DeleteProformaAsync(string pInvoice);
    }
}
