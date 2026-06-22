using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services
{
    public class ForeignCustomerService : IForeignCustomerService
    {
        private readonly IForeignCustomerDataAccess _dac;

        public ForeignCustomerService(IForeignCustomerDataAccess dac)
        {
            _dac = dac;
        }

        // ------------------------------------------------------------------ //
        //  Lookups — static or DB-driven
        // ------------------------------------------------------------------ //

        public Task<List<GenericDropDownModel>> GetCountriesAsync() =>
            Task.FromResult(new List<GenericDropDownModel>());   // populated at runtime via IDBHelperService in the component

        public Task<List<GenericDropDownModel>> GetCurrenciesAsync() =>
            Task.FromResult(new List<GenericDropDownModel>());

        public Task<List<GenericDropDownModel>> GetPaymentTermsAsync() =>
            Task.FromResult(new List<GenericDropDownModel>());

        public Task<List<GenericDropDownModel>> GetTradeTermsAsync() =>
            Task.FromResult(new List<GenericDropDownModel>());

        public Task<List<GenericDropDownModel>> GetFinishingQualitiesAsync() =>
            Task.FromResult(new List<GenericDropDownModel>());

        public Task<List<GenericDropDownModel>> GetShippingMethodsAsync() =>
            Task.FromResult(new List<GenericDropDownModel>());

        public Task<List<GenericDropDownModel>> GetCustomerSourcesAsync()
        {
            var items = new List<GenericDropDownModel>
            {
                new() { DropDownValue_ID = "1", DropDownValue_Description = "Exhibition" },
                new() { DropDownValue_ID = "2", DropDownValue_Description = "Web Promotion" },
                new() { DropDownValue_ID = "3", DropDownValue_Description = "Advertisement" },
                new() { DropDownValue_ID = "4", DropDownValue_Description = "Customer Reference" },
                new() { DropDownValue_ID = "5", DropDownValue_Description = "Personal Visit" },
                new() { DropDownValue_ID = "6", DropDownValue_Description = "Other" }
            };
            return Task.FromResult(items);
        }

        public Task<List<GenericDropDownModel>> GetEventRepetitionsAsync()
        {
            var items = new List<GenericDropDownModel>
            {
                new() { DropDownValue_ID = "0", DropDownValue_Description = "No Repeat" },
                new() { DropDownValue_ID = "1", DropDownValue_Description = "Every Year" }
            };
            return Task.FromResult(items);
        }

        // ------------------------------------------------------------------ //
        //  Customer CRUD
        // ------------------------------------------------------------------ //

        public async Task<ForeignCustomerViewModel?> GetCustomerAsync(string custId) =>
            await _dac.GetCustomerAsync(custId);

        public async Task<string> SaveCustomerAsync(
            ForeignCustomerViewModel model,
            string custId,
            string userName,
            string computerName,
            List<int> removedBankIds,
            List<int> removedAddressIds,
            List<int> removedEventIds,
            List<int> removedInvoiceToIds)
        {
            bool isNew = string.IsNullOrEmpty(custId);

            if (isNew)
            {
                custId = await _dac.InsertCustomerAsync(model, userName, computerName);
            }
            else
            {
                await _dac.UpdateCustomerAsync(model, custId);
            }

            await _dac.SaveBanksAsync(model, model.CustCode, model.Country, removedBankIds);
            await _dac.SaveShippingAddressesAsync(model, model.CustCode, model.Country, removedAddressIds);
            await _dac.SaveEventsAsync(model, model.CustCode, model.Country, removedEventIds);
            await _dac.SaveInvoiceToAsync(model, model.CustCode, model.Country, removedInvoiceToIds);

            return custId;
        }
    }
}
