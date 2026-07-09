using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class ProformaInvoiceService : IProformaInvoiceService
    {
        private readonly IProformaInvoiceDataAccess _dac;

        public ProformaInvoiceService(IProformaInvoiceDataAccess dac) => _dac = dac;

        public Task<List<GenericDropDownModel>> GetCompaniesAsync() => _dac.GetCompaniesAsync();
        public Task<List<CustomerLookupModel>> GetCustomersAsync() => _dac.GetCustomersAsync();
        public Task<List<string>> GetCountriesForCustomerAsync(string custCode) => _dac.GetCountriesForCustomerAsync(custCode);
        public Task<string> GetCurrencyForCustomerAsync(string custCode, string country) => _dac.GetCurrencyForCustomerAsync(custCode, country);
        public Task<List<GenericDropDownModel>> GetTradeTermsAsync() => _dac.GetTradeTermsAsync();
        public Task<List<GenericDropDownModel>> GetPaymentTermsAsync() => _dac.GetPaymentTermsAsync();
        public Task<List<GenericDropDownModel>> GetShippingMethodsAsync() => _dac.GetShippingMethodsAsync();
        public Task<List<AgentLookupModel>> GetShippingAgentsAsync() => _dac.GetShippingAgentsAsync();
        public Task<List<AgentLookupModel>> GetInsuranceAgentsAsync() => _dac.GetInsuranceAgentsAsync();
        public Task<List<CustomerBankModel>> GetCustomerBanksAsync(string custCode, string country) => _dac.GetCustomerBanksAsync(custCode, country);
        public Task<List<CompanyBankAccountModel>> GetCompanyBankAccountsAsync(int companyId) => _dac.GetCompanyBankAccountsAsync(companyId);
        public Task<List<ShippingAddressPresetModel>> GetShippingAddressPresetsAsync(string custCode, string country) => _dac.GetShippingAddressPresetsAsync(custCode, country);
        public Task<List<InvoiceToPresetModel>> GetInvoiceToPresetsAsync(string custCode, string country) => _dac.GetInvoiceToPresetsAsync(custCode, country);
        public Task<List<string>> GetOrdersForProformaAsync(int companyId, string custCode, string country) => _dac.GetOrdersForProformaAsync(companyId, custCode, country);
        public Task<List<OrderItemForProformaModel>> GetOrderItemsForProformaAsync(string orderNo) => _dac.GetOrderItemsForProformaAsync(orderNo);
        public Task<List<CityModel>> GetCitiesAsync(string countryName) => _dac.GetCitiesAsync(countryName);
        public Task<List<PortModel>> GetPortsAsync(int cityId) => _dac.GetPortsAsync(cityId);
        public Task<(int CityID, string CountryName)> GetPortCityAndCountryAsync(int portId) => _dac.GetPortCityAndCountryAsync(portId);
        public Task<(bool AcceptsExtraQty, string DefaultShipMethod, string PaymentTerms)> GetCustomerCountryDefaultsAsync(string custCode, string country) => _dac.GetCustomerCountryDefaultsAsync(custCode, country);
        public Task<string> GetNextProformaNumberAsync() => _dac.GetNextProformaNumberAsync();
        public Task<ProformaInvoiceViewModel?> GetProformaInvoiceAsync(string pInvoice) => _dac.GetProformaInvoiceAsync(pInvoice);
        public Task<bool> SaveProformaInvoiceAsync(ProformaInvoiceViewModel dto) => _dac.SaveProformaInvoiceAsync(dto);
        public Task<bool> DeleteProformaInvoiceAsync(string pInvoice) => _dac.DeleteProformaInvoiceAsync(pInvoice);
        public Task<double> CalculateSteelWeightAsync(string pInvoice) => _dac.CalculateSteelWeightAsync(pInvoice);
    }
}
