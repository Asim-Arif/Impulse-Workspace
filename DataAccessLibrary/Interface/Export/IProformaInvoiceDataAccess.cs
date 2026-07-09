using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;

namespace DataAccessLibrary.Interface.Export
{
    public interface IProformaInvoiceDataAccess
    {
        // Lookup data
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
        Task<double> GetCustomerDefaultPaymentTermsAsync(string custCode, string country);
        Task<string> GetDefaultShipMethodAsync(string custCode, string country);

        // Proforma # generation
        Task<string> GetNextProformaNumberAsync();

        // Load / save
        Task<ProformaInvoiceViewModel?> GetProformaInvoiceAsync(string pInvoice);
        Task<bool> SaveProformaInvoiceAsync(ProformaInvoiceViewModel dto);
        Task<bool> DeleteProformaInvoiceAsync(string pInvoice);
        Task<double> CalculateSteelWeightAsync(string pInvoice);

        // ForeignCustomer defaults on country select
        Task<(bool AcceptsExtraQty, string DefaultShipMethod, string PaymentTerms)> GetCustomerCountryDefaultsAsync(string custCode, string country);
    }
}
