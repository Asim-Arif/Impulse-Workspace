using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Export
{
    public class ProformaListService : IProformaListService
    {
        private readonly IProformaListDataAccess _dataAccess;

        public ProformaListService(IProformaListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<IEnumerable<ProformaListModel>> GetProformaListAsync(DateTime fromDate, DateTime toDate, string custCode, long companyId, string country)
        {
            return await _dataAccess.GetProformaListAsync(fromDate, toDate, custCode, companyId, country);
        }

        public async Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync()
        {
            return await _dataAccess.GetCustomersAsync();
        }

        public async Task<IEnumerable<CompanyLookupModel>> GetCompaniesAsync()
        {
            return await _dataAccess.GetCompaniesAsync();
        }

        public async Task<IEnumerable<string>> GetCountriesAsync(string custCode)
        {
            return await _dataAccess.GetCountriesAsync(custCode);
        }

        public async Task<bool> DeleteProformaAsync(string pInvoice)
        {
            return await _dataAccess.DeleteProformaAsync(pInvoice);
        }
    }
}
