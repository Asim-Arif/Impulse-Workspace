using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface ICustomerOrderDataAccess
    {
        // Lookup / Default loading methods
        Task<List<GenericDropDownModel>> GetCompaniesAsync();
        Task<List<CustomerLookupModel>> GetCustomersAsync();
        Task<List<GenericDropDownModel>> GetTradeTermsAsync();
        Task<List<GenericDropDownModel>> GetPaymentTermsAsync();
        Task<List<GenericDropDownModel>> GetShippingMethodsAsync();
        Task<List<GenericDropDownModel>> GetQualitiesAsync();
        Task<List<GenericDropDownModel>> GetFinishingQualitiesAsync();
        Task<List<string>> GetCountriesForCustomerAsync(string custCode);
        Task<string> GetCurrencyForCustomerAsync(string custCode, string country);
        Task<ForeignCustomerViewModel?> GetCustomerDefaultsAsync(string custCode, string country);

        // Catalog & Item queries
        Task<List<CustomerOrderItemViewModel>> GetCustomerCatalogItemsAsync(string custCode, string country);
        Task<decimal> GetItemCustomRateAsync(string custCode, string country, string compItemId, string finishingField);
        Task<decimal> GetItemWeightAsync(string itemId);

        // Quotation loading
        Task<List<string>> GetQuotationNumbersAsync(string custCode, string country);
        Task<List<CustomerOrderItemViewModel>> GetQuotationItemsAsync(string quotationNo);

        // Order CRUD operations
        Task<CustomerOrderViewModel?> GetOrderAsync(string orderNo);
        Task<bool> SaveOrderAsync(CustomerOrderViewModel order, List<int> deletedItemIds);
        Task<string> GetNextInternalRefNoAsync(int companyId);
        Task<bool> IsOrderNoExistsAsync(string orderNo);
        Task<bool> IsOrderUsedInProformaAsync(string orderNo);

        // Excel Import operations
        Task<bool> ClearExcelImportTableAsync();
        Task<bool> InsertExcelImportRowAsync(string itemId, string itemName, decimal qty, string unit, decimal rate, decimal weight, string steelType, string udms);
        Task<bool> ProcessExcelImportAsync(string custCode, string country);
        Task<List<CustomerOrderItemViewModel>> GetExcelImportErrorsAsync(bool compCatalogMissing);
        Task<List<CustomerOrderItemViewModel>> GetExcelImportOkAsync(string custCode, string country);
    }
}
