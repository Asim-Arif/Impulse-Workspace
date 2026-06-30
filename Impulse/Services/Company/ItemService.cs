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
        public Task<bool> SaveItemAsync(ItemDto item, bool isAdd)        => _dataAccess.SaveItemAsync(item, isAdd);
    }
}
