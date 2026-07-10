using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;

namespace DataAccessLibrary.Interface.Export
{
    public interface ICustomInvoiceDataAccess
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
        Task<List<GenericDropDownModel>> GetCustomerItemsAsync(string custCode, string country);
        Task<List<GenericDropDownModel>> GetProformasByItemAsync(string itemCode, string custCode, string country);
        Task<List<string>> GetProformasForInvoiceAsync(string custCode, string country);
        Task<List<CustomPListModel>> GetItemsForProformaAsync(string pInvoice);
        Task<List<DispatchListLookupModel>> GetAvailableDispatchListsAsync(string custCode);
        Task<List<DispatchCartonModel>> GetDispatchListCartonsAsync(long dispatchListEntryId);
        Task<List<CustomPListModel>> GetProformaItemsForDispatchMappingAsync(string orderNo, string itemCode);
        Task<CustomInvoiceViewModel?> GetProformaHeaderAsync(string pInvoice);
        Task<List<CustomInvoiceOtherChargeModel>> GetProformaOtherChargesAsync(string pInvoice);
        Task<List<CityModel>> GetCitiesAsync(string countryName);
        Task<List<PortModel>> GetPortsAsync(int cityId);
        Task<(int CityID, string CountryName)> GetPortCityAndCountryAsync(int portId);
        Task<List<AdvancePaymentModel>> GetAdvancePaymentsAsync(string custCode);

        // Custom Invoice # generation
        Task<string> GetNextCustomInvoiceNumberAsync(string strCompany);

        // Load / save
        Task<CustomInvoiceViewModel?> GetCustomInvoiceAsync(string customInvoiceNo);
        Task<bool> SaveCustomInvoiceAsync(CustomInvoiceViewModel dto);
        Task<bool> DeleteCustomInvoiceAsync(string customInvoiceNo);
    }
}
