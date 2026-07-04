using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Company
{
    public interface ICustomerCatalogDataAccess
    {
        Task<List<CustomerLookupModel>> GetCustomersAsync();
        Task<List<string>> GetCountriesForCustomerAsync(string custCode);
        Task<double> GetMaxDiscountAsync(string custCode, string country);
        Task<List<ItemGroupLookupModel>> GetItemGroupsAsync();
        
        Task<List<CustomerCatalogItemViewModel>> GetCatalogItemsAsync(string custCode, string country, int? groupID);
        Task<List<ItemLookupModel>> GetUnassignedItemsAsync(string custCode, string country);
        
        Task<int> AssignItemAsync(string custCode, string country, string itemId);
        Task<int> AssignAllItemsAsync(string custCode, string country);
        
        Task<string?> ItemIsInOrderAsync(string compItemId, string custCode, string country);
        Task<int> UnassignItemAsync(string custCode, string country, string compItemId);
        Task<int> UnassignAllItemsAsync(string custCode, string country);
        
        Task<int> UpdateRatesAsync(string custCode, double percentage, bool isIncrease);
        Task<bool> UpdateCatalogItemDetailAsync(CustomerCatalogItemViewModel model);
        
        Task<List<CustomerCatalogRevisionViewModel>> GetRevisionsAsync(int entryId);
        Task<byte[]> GetItemPicAsync(string compItemId);

        // Edit modal lookups
        Task<List<string>> GetUnitsAsync();
        Task<List<FinQualityLookupModel>> GetFinishingQualitiesAsync();
        Task<List<string>> GetMDGroupsAsync();
        Task<bool> UploadArtworkPDFAsync(int entryId, byte[] pdfBytes);
        Task<string> GetPageNoAsync(int entryId);
    }
}
