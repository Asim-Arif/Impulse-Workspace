using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface ICommercialCoveringService
    {
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<CustomerCountryModel>> GetCountriesForCustomerAsync(string custCode);
        Task<List<string>> GetInvoicesWithLettersAsync(string custCode, string country);
        Task<List<string>> GetInvoicesWithoutLettersAsync(string custCode, string country);
        Task<List<CoveringLetterItemViewModel>> GetCoveringLetterItemsAsync(string customInvoice);
        Task<bool> SaveCoveringLetterItemsAsync(string customInvoice, List<CoveringLetterItemViewModel> items);
    }
}
