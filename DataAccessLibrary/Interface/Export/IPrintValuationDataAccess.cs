using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface IPrintValuationDataAccess
    {
        Task<List<CustomerLookupModel>> GetCustomersAsync();
        Task<List<CustomerCountryModel>> GetCountriesForCustomerAsync(string custCode);
        Task<List<string>> GetInvoicesAsync(string custCode, string country);
    }
}
