using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Company
{
    public class CustomerCatalogService : ICustomerCatalogService
    {
        private readonly ICustomerCatalogDataAccess _dataAccess;

        public CustomerCatalogService(ICustomerCatalogDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<CustomerLookupModel>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        
        public Task<List<string>> GetCountriesForCustomerAsync(string custCode) => _dataAccess.GetCountriesForCustomerAsync(custCode);
        
        public Task<double> GetMaxDiscountAsync(string custCode, string country) => _dataAccess.GetMaxDiscountAsync(custCode, country);
        
        public Task<List<ItemGroupLookupModel>> GetItemGroupsAsync() => _dataAccess.GetItemGroupsAsync();
        
        public Task<List<CustomerCatalogItemViewModel>> GetCatalogItemsAsync(string custCode, string country, int? groupID) => 
            _dataAccess.GetCatalogItemsAsync(custCode, country, groupID);
        
        public Task<List<ItemLookupModel>> GetUnassignedItemsAsync(string custCode, string country) => 
            _dataAccess.GetUnassignedItemsAsync(custCode, country);
        
        public Task<int> AssignItemAsync(string custCode, string country, string itemId) => 
            _dataAccess.AssignItemAsync(custCode, country, itemId);
        
        public Task<int> AssignAllItemsAsync(string custCode, string country) => 
            _dataAccess.AssignAllItemsAsync(custCode, country);
        
        public Task<string?> ItemIsInOrderAsync(string compItemId, string custCode, string country) => 
            _dataAccess.ItemIsInOrderAsync(compItemId, custCode, country);
        
        public Task<int> UnassignItemAsync(string custCode, string country, string compItemId) => 
            _dataAccess.UnassignItemAsync(custCode, country, compItemId);
        
        public Task<int> UnassignAllItemsAsync(string custCode, string country) => 
            _dataAccess.UnassignAllItemsAsync(custCode, country);
        
        public Task<int> UpdateRatesAsync(string custCode, double percentage, bool isIncrease) => 
            _dataAccess.UpdateRatesAsync(custCode, percentage, isIncrease);
        
        public Task<bool> UpdateCatalogItemDetailAsync(CustomerCatalogItemViewModel model) => 
            _dataAccess.UpdateCatalogItemDetailAsync(model);
        
        public Task<List<CustomerCatalogRevisionViewModel>> GetRevisionsAsync(int entryId) => 
            _dataAccess.GetRevisionsAsync(entryId);
        
        public Task<byte[]> GetItemPicAsync(string compItemId) => 
            _dataAccess.GetItemPicAsync(compItemId);

        // Edit modal lookups
        public Task<List<string>> GetUnitsAsync() => _dataAccess.GetUnitsAsync();
        public Task<List<FinQualityLookupModel>> GetFinishingQualitiesAsync() => _dataAccess.GetFinishingQualitiesAsync();
        public Task<List<string>> GetMDGroupsAsync() => _dataAccess.GetMDGroupsAsync();
        public Task<bool> UploadArtworkPDFAsync(int entryId, byte[] pdfBytes) => _dataAccess.UploadArtworkPDFAsync(entryId, pdfBytes);
        public Task<string> GetPageNoAsync(int entryId) => _dataAccess.GetPageNoAsync(entryId);
    }
}
