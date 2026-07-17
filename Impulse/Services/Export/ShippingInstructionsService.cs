using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class ShippingInstructionsService : IShippingInstructionsService
    {
        private readonly IShippingInstructionsDataAccess _dataAccess;

        public ShippingInstructionsService(IShippingInstructionsDataAccess dataAccess)
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

        public async Task<List<string>> GetInvoicesWithInstructionsAsync(string custCode, string country)
        {
            return await _dataAccess.GetInvoicesWithInstructionsAsync(custCode, country);
        }

        public async Task<List<string>> GetInvoicesWithoutInstructionsAsync(string custCode, string country)
        {
            return await _dataAccess.GetInvoicesWithoutInstructionsAsync(custCode, country);
        }

        public async Task<ShippingInstructionsViewModel?> GetShippingInstructionsAsync(string customInvoice)
        {
            return await _dataAccess.GetShippingInstructionsAsync(customInvoice);
        }

        public async Task<(int TotalCartons, string InvoiceOf)> GetTotalCartonsAndInvoiceOfAsync(string customInvoice)
        {
            return await _dataAccess.GetTotalCartonsAndInvoiceOfAsync(customInvoice);
        }

        public async Task<double> GetDefaultBoxSizeMeasurementAsync(int totalCartons)
        {
            return await _dataAccess.GetDefaultBoxSizeMeasurementAsync(totalCartons);
        }

        public async Task<bool> SaveShippingInstructionsAsync(ShippingInstructionsViewModel model)
        {
            return await _dataAccess.SaveShippingInstructionsAsync(model);
        }
    }
}
