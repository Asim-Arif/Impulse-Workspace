using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface IShippingInstructionsDataAccess
    {
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<string>> GetCountriesForCustomerAsync(string custCode);
        Task<List<string>> GetInvoicesWithInstructionsAsync(string custCode, string country);
        Task<List<string>> GetInvoicesWithoutInstructionsAsync(string custCode, string country);
        Task<ShippingInstructionsViewModel?> GetShippingInstructionsAsync(string customInvoice);
        Task<(int TotalCartons, string InvoiceOf)> GetTotalCartonsAndInvoiceOfAsync(string customInvoice);
        Task<double> GetDefaultBoxSizeMeasurementAsync(int totalCartons);
        Task<bool> SaveShippingInstructionsAsync(ShippingInstructionsViewModel model);
    }
}
