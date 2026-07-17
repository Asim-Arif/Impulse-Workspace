using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class PackingLabelsService : IPackingLabelsService
    {
        private readonly IPackingLabelsDataAccess _dataAccess;

        public PackingLabelsService(IPackingLabelsDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            return await _dataAccess.GetActiveCustomersAsync();
        }

        public async Task<List<string>> GetCountriesForCustomerAsync(string custCode)
        {
            return await _dataAccess.GetCountriesForCustomerAsync(custCode);
        }

        public async Task<CustomerLabelSettingsModel?> GetLabelSettingsAsync(string custCode, string country)
        {
            return await _dataAccess.GetLabelSettingsAsync(custCode, country);
        }

        public async Task<List<string>> GetCustomInvoicesAsync(string custCode, string country)
        {
            return await _dataAccess.GetCustomInvoicesAsync(custCode, country);
        }

        public async Task<int> GetTotalCartonsAsync(string customInvoice)
        {
            return await _dataAccess.GetTotalCartonsAsync(customInvoice);
        }

        public async Task GeneratePrintPackingLabelsAsync(string customInvoice)
        {
            await _dataAccess.GeneratePrintPackingLabelsAsync(customInvoice);
        }
    }
}
