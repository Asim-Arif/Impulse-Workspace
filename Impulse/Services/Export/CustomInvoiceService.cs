using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;

namespace Impulse.Services.Export
{
    public class CustomInvoiceService : ICustomInvoiceService
    {
        private readonly ICustomInvoiceDataAccess _dataAccess;

        public CustomInvoiceService(ICustomInvoiceDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<GenericDropDownModel>> GetCompaniesAsync() => _dataAccess.GetCompaniesAsync();
        public Task<List<CustomerLookupModel>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        public Task<List<string>> GetCountriesForCustomerAsync(string custCode) => _dataAccess.GetCountriesForCustomerAsync(custCode);
        public Task<string> GetCurrencyForCustomerAsync(string custCode, string country) => _dataAccess.GetCurrencyForCustomerAsync(custCode, country);
        public Task<List<GenericDropDownModel>> GetTradeTermsAsync() => _dataAccess.GetTradeTermsAsync();
        public Task<List<GenericDropDownModel>> GetPaymentTermsAsync() => _dataAccess.GetPaymentTermsAsync();
        public Task<List<GenericDropDownModel>> GetShippingMethodsAsync() => _dataAccess.GetShippingMethodsAsync();
        public Task<List<AgentLookupModel>> GetShippingAgentsAsync() => _dataAccess.GetShippingAgentsAsync();
        public Task<List<AgentLookupModel>> GetInsuranceAgentsAsync() => _dataAccess.GetInsuranceAgentsAsync();
        public Task<List<CustomerBankModel>> GetCustomerBanksAsync(string custCode, string country) => _dataAccess.GetCustomerBanksAsync(custCode, country);
        public Task<List<CompanyBankAccountModel>> GetCompanyBankAccountsAsync(int companyId) => _dataAccess.GetCompanyBankAccountsAsync(companyId);
        public Task<List<ShippingAddressPresetModel>> GetShippingAddressPresetsAsync(string custCode, string country) => _dataAccess.GetShippingAddressPresetsAsync(custCode, country);
        public Task<List<InvoiceToPresetModel>> GetInvoiceToPresetsAsync(string custCode, string country) => _dataAccess.GetInvoiceToPresetsAsync(custCode, country);
        public Task<List<GenericDropDownModel>> GetCustomerItemsAsync(string custCode, string country) => _dataAccess.GetCustomerItemsAsync(custCode, country);
        public Task<List<GenericDropDownModel>> GetProformasByItemAsync(string itemCode, string custCode, string country) => _dataAccess.GetProformasByItemAsync(itemCode, custCode, country);
        public Task<List<string>> GetProformasForInvoiceAsync(string custCode, string country) => _dataAccess.GetProformasForInvoiceAsync(custCode, country);
        public Task<List<CustomPListModel>> GetItemsForProformaAsync(string pInvoice) => _dataAccess.GetItemsForProformaAsync(pInvoice);
        public Task<List<DispatchListLookupModel>> GetAvailableDispatchListsAsync(string custCode) => _dataAccess.GetAvailableDispatchListsAsync(custCode);
        public Task<List<DispatchCartonModel>> GetDispatchListCartonsAsync(long dispatchListEntryId) => _dataAccess.GetDispatchListCartonsAsync(dispatchListEntryId);
        public Task<List<CustomPListModel>> GetProformaItemsForDispatchMappingAsync(string orderNo, string itemCode) => _dataAccess.GetProformaItemsForDispatchMappingAsync(orderNo, itemCode);
        public Task<CustomInvoiceViewModel?> GetProformaHeaderAsync(string pInvoice) => _dataAccess.GetProformaHeaderAsync(pInvoice);
        public Task<List<CustomInvoiceOtherChargeModel>> GetProformaOtherChargesAsync(string pInvoice) => _dataAccess.GetProformaOtherChargesAsync(pInvoice);
        public Task<List<CityModel>> GetCitiesAsync(string countryName) => _dataAccess.GetCitiesAsync(countryName);
        public Task<List<PortModel>> GetPortsAsync(int cityId) => _dataAccess.GetPortsAsync(cityId);
        public Task<(int CityID, string CountryName)> GetPortCityAndCountryAsync(int portId) => _dataAccess.GetPortCityAndCountryAsync(portId);
        public Task<List<AdvancePaymentModel>> GetAdvancePaymentsAsync(string custCode) => _dataAccess.GetAdvancePaymentsAsync(custCode);

        public Task<string> GetNextCustomInvoiceNumberAsync(string strCompany) => _dataAccess.GetNextCustomInvoiceNumberAsync(strCompany);

        public Task<CustomInvoiceViewModel?> GetCustomInvoiceAsync(string customInvoiceNo) => _dataAccess.GetCustomInvoiceAsync(customInvoiceNo);
        public Task<bool> SaveCustomInvoiceAsync(CustomInvoiceViewModel dto) => _dataAccess.SaveCustomInvoiceAsync(dto);
        public Task<bool> DeleteCustomInvoiceAsync(string customInvoiceNo) => _dataAccess.DeleteCustomInvoiceAsync(customInvoiceNo);
    }
}
