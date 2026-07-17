using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class PrintValuationService : IPrintValuationService
    {
        private readonly IPrintValuationDataAccess _dataAccess;

        public PrintValuationService(IPrintValuationDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            return await _dataAccess.GetCustomersAsync();
        }

        public async Task<List<CustomerCountryModel>> GetCountriesForCustomerAsync(string custCode)
        {
            return await _dataAccess.GetCountriesForCustomerAsync(custCode);
        }

        public async Task<List<string>> GetInvoicesAsync(string custCode, string country)
        {
            return await _dataAccess.GetInvoicesAsync(custCode, country);
        }
    }
}
