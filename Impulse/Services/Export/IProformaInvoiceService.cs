using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface IProformaInvoiceService
    {
        Task<List<GenericDropDownModel>> GetCompaniesAsync();
        Task<List<CustomerLookupModel>> GetCustomersAsync();
        Task<List<string>> GetCountriesForCustomerAsync(string custCode);
        Task<string> GetCurrencyForCustomerAsync(string custCode, string country);
        Task<List<GenericDropDownModel>> GetTradeTermsAsync();
        Task<List<GenericDropDownModel>> GetPaymentTermsAsync();
        Task<List<GenericDropDownModel>> GetShippingMethodsAsync();
        Task<List<AgentLookupModel>> GetShippingAgentsAsync();
        Task<List<AgentLookupModel>> GetInsuranceAgentsAsync();
        Task<List<CustomerBankModel>> GetCustomerBanksAsync(string custCode, string country);
        Task<List<CompanyBankAccountModel>> GetCompanyBankAccountsAsync(int companyId);
        Task<List<ShippingAddressPresetModel>> GetShippingAddressPresetsAsync(string custCode, string country);
        Task<List<InvoiceToPresetModel>> GetInvoiceToPresetsAsync(string custCode, string country);
        Task<List<string>> GetOrdersForProformaAsync(int companyId, string custCode, string country);
        Task<List<OrderItemForProformaModel>> GetOrderItemsForProformaAsync(string orderNo);
        Task<List<CityModel>> GetCitiesAsync(string countryName);
        Task<List<PortModel>> GetPortsAsync(int cityId);
        Task<(int CityID, string CountryName)> GetPortCityAndCountryAsync(int portId);
        Task<(bool AcceptsExtraQty, string DefaultShipMethod, string PaymentTerms)> GetCustomerCountryDefaultsAsync(string custCode, string country);
        Task<string> GetNextProformaNumberAsync();
        Task<ProformaInvoiceViewModel?> GetProformaInvoiceAsync(string pInvoice);
        Task<bool> SaveProformaInvoiceAsync(ProformaInvoiceViewModel dto);
        Task<bool> DeleteProformaInvoiceAsync(string pInvoice);
        Task<double> CalculateSteelWeightAsync(string pInvoice);
    }
}
