using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface IPackingLabelsDataAccess
    {
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<string>> GetCountriesForCustomerAsync(string custCode);
        Task<CustomerLabelSettingsModel?> GetLabelSettingsAsync(string custCode, string country);
        Task<List<string>> GetCustomInvoicesAsync(string custCode, string country);
        Task<int> GetTotalCartonsAsync(string customInvoice);
        Task GeneratePrintPackingLabelsAsync(string customInvoice);
    }
}
