using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface IPackingLabelsService
    {
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<string>> GetCountriesForCustomerAsync(string custCode);
        Task<CustomerLabelSettingsModel?> GetLabelSettingsAsync(string custCode, string country);
        Task<List<string>> GetCustomInvoicesAsync(string custCode, string country);
        Task<int> GetTotalCartonsAsync(string customInvoice);
        Task GeneratePrintPackingLabelsAsync(string customInvoice);
    }
}
