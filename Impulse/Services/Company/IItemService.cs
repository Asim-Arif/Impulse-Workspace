using DataAccessLibrary.Models.ViewModels.Company;

namespace Impulse.Services.Company
{
    public interface IItemService
    {
        Task<List<ItemCategoryModel>> GetCategoriesAsync();
        Task<List<ItemGroupModel>> GetItemGroupsAsync();
        Task<List<ItemMainGroupModel>> GetMainGroupsAsync();
        Task<List<UnitModel>> GetUnitsAsync();
        Task<List<ItemTypeModel>> GetItemTypesAsync();
        Task<ItemDto?> GetItemByIdAsync(string itemId);
        Task<bool> ItemIdExistsAsync(string itemId);
        Task<List<ItemProcessModel>> GetItemProcessesAsync(string itemId);
        Task<List<ItemCatalogRefModel>> GetItemCatalogRefsAsync(string itemId);
        Task<bool> SaveItemAsync(ItemDto item, bool isAdd);
    }
}
