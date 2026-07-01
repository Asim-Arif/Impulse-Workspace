using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class CustomerOrderService : ICustomerOrderService
    {
        private readonly ICustomerOrderDataAccess _dac;

        public CustomerOrderService(ICustomerOrderDataAccess dac)
        {
            _dac = dac;
        }

        public async Task<List<GenericDropDownModel>> GetCompaniesAsync() => await _dac.GetCompaniesAsync();
        public async Task<List<CustomerLookupModel>> GetCustomersAsync() => await _dac.GetCustomersAsync();
        public async Task<List<GenericDropDownModel>> GetTradeTermsAsync() => await _dac.GetTradeTermsAsync();
        public async Task<List<GenericDropDownModel>> GetPaymentTermsAsync() => await _dac.GetPaymentTermsAsync();
        public async Task<List<GenericDropDownModel>> GetShippingMethodsAsync() => await _dac.GetShippingMethodsAsync();
        public async Task<List<GenericDropDownModel>> GetQualitiesAsync() => await _dac.GetQualitiesAsync();
        public async Task<List<GenericDropDownModel>> GetFinishingQualitiesAsync() => await _dac.GetFinishingQualitiesAsync();
        public async Task<List<string>> GetCountriesForCustomerAsync(string custCode) => await _dac.GetCountriesForCustomerAsync(custCode);
        public async Task<string> GetCurrencyForCustomerAsync(string custCode, string country) => await _dac.GetCurrencyForCustomerAsync(custCode, country);
        public async Task<ForeignCustomerViewModel?> GetCustomerDefaultsAsync(string custCode, string country) => await _dac.GetCustomerDefaultsAsync(custCode, country);

        public async Task<List<CustomerOrderItemViewModel>> GetCustomerCatalogItemsAsync(string custCode, string country) => await _dac.GetCustomerCatalogItemsAsync(custCode, country);
        public async Task<decimal> GetItemCustomRateAsync(string custCode, string country, string compItemId, string finishingField) => await _dac.GetItemCustomRateAsync(custCode, country, compItemId, finishingField);
        public async Task<decimal> GetItemWeightAsync(string itemId) => await _dac.GetItemWeightAsync(itemId);

        public async Task<List<string>> GetQuotationNumbersAsync(string custCode, string country) => await _dac.GetQuotationNumbersAsync(custCode, country);
        public async Task<List<CustomerOrderItemViewModel>> GetQuotationItemsAsync(string quotationNo) => await _dac.GetQuotationItemsAsync(quotationNo);

        public async Task<CustomerOrderViewModel?> GetOrderAsync(string orderNo) => await _dac.GetOrderAsync(orderNo);
        public async Task<bool> SaveOrderAsync(CustomerOrderViewModel order, List<int> deletedItemIds) => await _dac.SaveOrderAsync(order, deletedItemIds);
        public async Task<string> GetNextInternalRefNoAsync(int companyId) => await _dac.GetNextInternalRefNoAsync(companyId);
        public async Task<bool> IsOrderNoExistsAsync(string orderNo) => await _dac.IsOrderNoExistsAsync(orderNo);
        public async Task<bool> IsOrderUsedInProformaAsync(string orderNo) => await _dac.IsOrderUsedInProformaAsync(orderNo);
        public async Task<bool> DeleteCustomerOrderAsync(string orderNo) => await _dac.DeleteCustomerOrderAsync(orderNo);

        public async Task<List<CustomerOrderListItemModel>> GetOrderListAsync(DateTime dtFrom, DateTime dtTo, string custCode, string country, int companyRefID, string orderType, int statusFilter, bool filterByDeliveryDT, int viewType) =>
            await _dac.GetOrderListAsync(dtFrom, dtTo, custCode, country, companyRefID, orderType, statusFilter, filterByDeliveryDT, viewType);
        public async Task<List<string>> GetAllCountriesAsync() => await _dac.GetAllCountriesAsync();
        public async Task<bool> UpdateOrderFinalStatusAsync(string orderNo, string custCode, string country, int cancelledStatus, string remarks, string userName, string machineName) =>
            await _dac.UpdateOrderFinalStatusAsync(orderNo, custCode, country, cancelledStatus, remarks, userName, machineName);
        public async Task<bool> DeleteOrderFinalStatusAsync(string orderNo) => await _dac.DeleteOrderFinalStatusAsync(orderNo);

        public async Task<bool> ClearExcelImportTableAsync() => await _dac.ClearExcelImportTableAsync();
        public async Task<bool> InsertExcelImportRowAsync(string itemId, string itemName, decimal qty, string unit, decimal rate, decimal weight, string steelType, string udms) =>
            await _dac.InsertExcelImportRowAsync(itemId, itemName, qty, unit, rate, weight, steelType, udms);
        public async Task<bool> ProcessExcelImportAsync(string custCode, string country) => await _dac.ProcessExcelImportAsync(custCode, country);
        public async Task<List<CustomerOrderItemViewModel>> GetExcelImportErrorsAsync(bool compCatalogMissing) => await _dac.GetExcelImportErrorsAsync(compCatalogMissing);
        public async Task<List<CustomerOrderItemViewModel>> GetExcelImportOkAsync(string custCode, string country) => await _dac.GetExcelImportOkAsync(custCode, country);
    }
}
