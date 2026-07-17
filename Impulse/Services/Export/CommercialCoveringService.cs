using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class CommercialCoveringService : ICommercialCoveringService
    {
        private readonly ICommercialCoveringDataAccess _dataAccess;

        public CommercialCoveringService(ICommercialCoveringDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            return await _dataAccess.GetActiveCustomersAsync();
        }

        public async Task<List<CustomerCountryModel>> GetCountriesForCustomerAsync(string custCode)
        {
            return await _dataAccess.GetCountriesForCustomerAsync(custCode);
        }

        public async Task<List<string>> GetInvoicesWithLettersAsync(string custCode, string country)
        {
            return await _dataAccess.GetInvoicesWithLettersAsync(custCode, country);
        }

        public async Task<List<string>> GetInvoicesWithoutLettersAsync(string custCode, string country)
        {
            return await _dataAccess.GetInvoicesWithoutLettersAsync(custCode, country);
        }

        public async Task<List<CoveringLetterItemViewModel>> GetCoveringLetterItemsAsync(string customInvoice)
        {
            return await _dataAccess.GetCoveringLetterItemsAsync(customInvoice);
        }

        public async Task<bool> SaveCoveringLetterItemsAsync(string customInvoice, List<CoveringLetterItemViewModel> items)
        {
            return await _dataAccess.SaveCoveringLetterItemsAsync(customInvoice, items);
        }
    }
}
