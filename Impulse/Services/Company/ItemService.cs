using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;

namespace Impulse.Services.Company
{
    public class ItemService : IItemService
    {
        private readonly IItemDataAccess _dataAccess;

        public ItemService(IItemDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ItemCategoryModel>> GetCategoriesAsync()       => _dataAccess.GetCategoriesAsync();
        public Task<List<ItemGroupModel>> GetItemGroupsAsync()           => _dataAccess.GetItemGroupsAsync();
        public Task<List<ItemMainGroupModel>> GetMainGroupsAsync()       => _dataAccess.GetMainGroupsAsync();
        public Task<List<UnitModel>> GetUnitsAsync()                     => _dataAccess.GetUnitsAsync();
        public Task<List<ItemTypeModel>> GetItemTypesAsync()             => _dataAccess.GetItemTypesAsync();
        public Task<ItemDto?> GetItemByIdAsync(string itemId)            => _dataAccess.GetItemByIdAsync(itemId);
        public Task<bool> ItemIdExistsAsync(string itemId)               => _dataAccess.ItemIdExistsAsync(itemId);
        public Task<List<ItemProcessModel>> GetItemProcessesAsync(string itemId)      => _dataAccess.GetItemProcessesAsync(itemId);
        public Task<List<ItemCatalogRefModel>> GetItemCatalogRefsAsync(string itemId) => _dataAccess.GetItemCatalogRefsAsync(itemId);
        public Task<List<ItemRMComponentModel>> GetItemRMComponentsAsync(string itemId) => _dataAccess.GetItemRMComponentsAsync(itemId);
        public Task<List<ItemLookAlikeModel>> GetItemLookAlikesAsync(string itemId)   => _dataAccess.GetItemLookAlikesAsync(itemId);
        public Task<List<ItemSetDetailModel>> GetItemSetDetailsAsync(string itemId)   => _dataAccess.GetItemSetDetailsAsync(itemId);
        public Task<List<ProcessGroupLookupModel>> GetProcessGroupsLookupAsync()       => _dataAccess.GetProcessGroupsLookupAsync();
        public Task<List<ProcessItemLookupModel>> GetProcessesLookupAsync(int? groupId = null) => _dataAccess.GetProcessesLookupAsync(groupId);
        public Task<(int? GroupId, string? GroupName)> GetItemProcessFamilyAsync(string itemId) => _dataAccess.GetItemProcessFamilyAsync(itemId);
        public Task<List<ForeignCatalogLookupModel>> GetForeignCatalogsLookupAsync()  => _dataAccess.GetForeignCatalogsLookupAsync();
        public Task<List<MaterialLookupModel>> GetMaterialLookupsAsync()               => _dataAccess.GetMaterialLookupsAsync();
        public Task<List<ItemSearchLookupModel>> GetItemSearchLookupsAsync()           => _dataAccess.GetItemSearchLookupsAsync();
        public Task<List<ItemSetLookupModel>> GetItemSetLookupsAsync()                 => _dataAccess.GetItemSetLookupsAsync();
        public Task<bool> SaveItemAsync(ItemDto item, bool isAdd)        => _dataAccess.SaveItemAsync(item, isAdd);
    }
}
