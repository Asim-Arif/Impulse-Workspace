using DataAccessLibrary.Models.ViewModels.Company;

namespace DataAccessLibrary.Interface.Company
{
    public interface IItemDataAccess
    {
        // ── Lookups ──────────────────────────────────────────────────────────
        Task<List<ItemCategoryModel>> GetCategoriesAsync();
        Task<List<ItemGroupModel>> GetItemGroupsAsync();
        Task<List<ItemMainGroupModel>> GetMainGroupsAsync();
        Task<List<UnitModel>> GetUnitsAsync();
        Task<List<ItemTypeModel>> GetItemTypesAsync();

        // ── Read ─────────────────────────────────────────────────────────────
        Task<ItemDto?> GetItemByIdAsync(string itemId);
        Task<bool> ItemIdExistsAsync(string itemId);

        // ── Child collections ────────────────────────────────────────────────
        Task<List<ItemProcessModel>> GetItemProcessesAsync(string itemId);
        Task<List<ItemCatalogRefModel>> GetItemCatalogRefsAsync(string itemId);

        // ── Persist ──────────────────────────────────────────────────────────
        Task<bool> SaveItemAsync(ItemDto item, bool isAdd);
    }
}
