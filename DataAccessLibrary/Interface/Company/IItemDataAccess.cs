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
        Task<List<ItemRMComponentModel>> GetItemRMComponentsAsync(string itemId);
        Task<List<ItemLookAlikeModel>> GetItemLookAlikesAsync(string itemId);
        Task<List<ItemSetDetailModel>> GetItemSetDetailsAsync(string itemId);
        Task<List<ProcessGroupLookupModel>> GetProcessGroupsLookupAsync();
        Task<List<ProcessItemLookupModel>> GetProcessesLookupAsync(int? groupId = null);
        Task<(int? GroupId, string? GroupName)> GetItemProcessFamilyAsync(string itemId);
        Task<List<ForeignCatalogLookupModel>> GetForeignCatalogsLookupAsync();
        Task<List<MaterialLookupModel>> GetMaterialLookupsAsync();
        Task<List<ItemSearchLookupModel>> GetItemSearchLookupsAsync();
        Task<List<ItemSetLookupModel>> GetItemSetLookupsAsync();

        // ── Persist ──────────────────────────────────────────────────────────
        Task<bool> SaveItemAsync(ItemDto item, bool isAdd);
    }
}
