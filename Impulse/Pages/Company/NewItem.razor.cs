using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services.Company;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Company
{
    public partial class NewItem
    {
        // ── Route & Query parameters ─────────────────────────────────────────
        [Parameter]
        public string? ItemId { get; set; }

        [Parameter]
        public string? CopyFromId { get; set; }

        [Parameter]
        [SupplyParameterFromQuery(Name = "id")]
        public string? QueryItemId { get; set; }

        [Parameter]
        [SupplyParameterFromQuery(Name = "copyFrom")]
        public string? QueryCopyFromId { get; set; }

        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        public string? EffectiveItemId => !string.IsNullOrWhiteSpace(QueryItemId) ? QueryItemId : ItemId;
        public string? EffectiveCopyFromId => !string.IsNullOrWhiteSpace(QueryCopyFromId) ? QueryCopyFromId : CopyFromId;

        // ── Injected services ────────────────────────────────────────────────
        [Inject] public IItemService ItemService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        // ── State ────────────────────────────────────────────────────────────
        public ItemDto Item { get; set; } = new();

        public List<ItemCategoryModel> Categories { get; set; } = new();
        public List<ItemGroupModel> ItemGroups { get; set; } = new();
        public List<ItemMainGroupModel> MainGroups { get; set; } = new();
        public List<UnitModel> Units { get; set; } = new();
        public List<ItemTypeModel> ItemTypes { get; set; } = new();
        public List<ForeignCatalogLookupModel> ForeignCatalogs { get; set; } = new();

        public ItemCategoryModel? SelectedCategory { get; set; }
        public ItemGroupModel? SelectedItemGroup { get; set; }
        public ItemMainGroupModel? SelectedMainGroup { get; set; }

        public string ActiveTab { get; set; } = "GeneralInfo";
        public bool IsSaving { get; set; }
        public bool IsAdd => string.IsNullOrEmpty(EffectiveItemId);

        // ── Lifecycle ────────────────────────────────────────────────────────

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();

            var effectiveCopyFrom = EffectiveCopyFromId;
            var effectiveItemId = EffectiveItemId;

            if (!string.IsNullOrEmpty(effectiveCopyFrom) && effectiveCopyFrom.Contains('%'))
            {
                try { effectiveCopyFrom = Uri.UnescapeDataString(effectiveCopyFrom); } catch { }
            }
            if (!string.IsNullOrEmpty(effectiveItemId) && effectiveItemId.Contains('%'))
            {
                try { effectiveItemId = Uri.UnescapeDataString(effectiveItemId); } catch { }
            }

            if (!string.IsNullOrEmpty(effectiveCopyFrom))
            {
                var existing = await ItemService.GetItemByIdAsync(effectiveCopyFrom);
                if (existing != null)
                {
                    Item = existing;
                    Item.ItemID = string.Empty;
                    Item.ItemName = $"{existing.ItemName} (Copy)";
                    var copySourceId = existing.ItemID;
                    Item.Processes = await ItemService.GetItemProcessesAsync(copySourceId);
                    foreach (var p in Item.Processes)
                    {
                        p.EntryID = null;
                    }
                    Item.CatalogRefs = await ItemService.GetItemCatalogRefsAsync(copySourceId);
                    foreach (var cr in Item.CatalogRefs)
                    {
                        cr.EntryID = null;
                    }
                    Item.RMComponents = await ItemService.GetItemRMComponentsAsync(copySourceId);
                    foreach (var rc in Item.RMComponents)
                    {
                        rc.EntryID = null;
                    }
                    Item.LookAlikes = await ItemService.GetItemLookAlikesAsync(copySourceId);
                    foreach (var l in Item.LookAlikes)
                    {
                        l.EntryID = null;
                    }
                    Item.SetDetails = await ItemService.GetItemSetDetailsAsync(copySourceId);
                    foreach (var s in Item.SetDetails)
                    {
                        s.EntryID = null;
                    }

                    SelectedCategory = Categories.FirstOrDefault(c => c.CatID == Item.CatID);
                    SelectedItemGroup = ItemGroups.FirstOrDefault(g => g.ID == Item.GroupID);
                    SelectedMainGroup = MainGroups.FirstOrDefault(m => m.MainGroupID == Item.MainGroupID);

                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Info,
                        Summary = "Item Copied",
                        Detail = $"Details copied from '{copySourceId}'. Enter a new Item Code and save.",
                        Duration = 5000
                    });
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Not Found",
                        Detail = $"Source item '{effectiveCopyFrom}' was not found.",
                        Duration = 4000
                    });
                    NavigationManager.NavigateTo(!string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/company/items");
                }
            }
            else if (!IsAdd && !string.IsNullOrEmpty(effectiveItemId))
            {
                var existing = await ItemService.GetItemByIdAsync(effectiveItemId);
                if (existing != null)
                {
                    Item = existing;
                    var resolvedId = existing.ItemID;
                    Item.Processes    = await ItemService.GetItemProcessesAsync(resolvedId);
                    Item.CatalogRefs  = await ItemService.GetItemCatalogRefsAsync(resolvedId);
                    Item.RMComponents = await ItemService.GetItemRMComponentsAsync(resolvedId);
                    Item.LookAlikes   = await ItemService.GetItemLookAlikesAsync(resolvedId);
                    Item.SetDetails   = await ItemService.GetItemSetDetailsAsync(resolvedId);

                    SelectedCategory  = Categories.FirstOrDefault(c => c.CatID       == Item.CatID);
                    SelectedItemGroup = ItemGroups.FirstOrDefault(g => g.ID          == Item.GroupID);
                    SelectedMainGroup = MainGroups.FirstOrDefault(m => m.MainGroupID == Item.MainGroupID);
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary  = "Not Found",
                        Detail   = $"Item '{effectiveItemId}' was not found.",
                        Duration = 4000
                    });
                    NavigationManager.NavigateTo(!string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/company/items");
                }
            }
        }

        private async Task LoadLookups()
        {
            try
            {
                var cats   = ItemService.GetCategoriesAsync();
                var groups = ItemService.GetItemGroupsAsync();
                var mains  = ItemService.GetMainGroupsAsync();
                var units  = ItemService.GetUnitsAsync();
                var types  = ItemService.GetItemTypesAsync();
                var foreignCatalogs = ItemService.GetForeignCatalogsLookupAsync();

                await Task.WhenAll(cats, groups, mains, units, types, foreignCatalogs);

                Categories = await cats;
                ItemGroups = await groups;
                MainGroups = await mains;
                Units      = await units;
                ItemTypes  = await types;
                ForeignCatalogs = await foreignCatalogs;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary  = "Error",
                    Detail   = $"Failed to load lookup data: {ex.Message}",
                    Duration = 4000
                });
            }
        }

        // ── Typeahead search methods ──────────────────────────────────────────

        public Task<IEnumerable<ItemCategoryModel>> SearchCategories(string text) =>
            Task.FromResult(Categories.Where(c =>
                c.Description.Contains(text, StringComparison.OrdinalIgnoreCase)));

        public Task<IEnumerable<ItemGroupModel>> SearchItemGroups(string text) =>
            Task.FromResult(ItemGroups.Where(g =>
                g.Description.Contains(text, StringComparison.OrdinalIgnoreCase)));

        public Task<IEnumerable<ItemMainGroupModel>> SearchMainGroups(string text) =>
            Task.FromResult(MainGroups.Where(m =>
                m.MainGroupName.Contains(text, StringComparison.OrdinalIgnoreCase)));

        // ── Selection handlers ────────────────────────────────────────────────

        public void OnCategorySelected(ItemCategoryModel? cat)
        {
            SelectedCategory = cat;
            Item.CatID = cat?.CatID;
        }

        public void OnGroupSelected(ItemGroupModel? grp)
        {
            SelectedItemGroup = grp;
            Item.GroupID = grp?.ID;
        }

        public void OnMainGroupSelected(ItemMainGroupModel? mg)
        {
            SelectedMainGroup = mg;
            Item.MainGroupID = mg?.MainGroupID;
        }

        // ── Tab navigation ────────────────────────────────────────────────────

        public void ChangeTab(string tab) => ActiveTab = tab;

        public void OnItemTypeChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int val))
            {
                Item.ItemType = val;
                if (Item.ItemType != 2 && ActiveTab == "SetDetail")
                {
                    ActiveTab = "GeneralInfo";
                }
            }
        }

        // ── Save / Cancel ─────────────────────────────────────────────────────

        public async Task HandleValidSubmit()
        {
            if (string.IsNullOrWhiteSpace(Item.ItemID))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary  = "Validation",
                    Detail   = "Item Code is required.",
                    Duration = 4000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(Item.ItemName))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary  = "Validation",
                    Detail   = "Item Name is required.",
                    Duration = 4000
                });
                return;
            }

            if (!Item.GroupID.HasValue)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary  = "Validation",
                    Detail   = "Item Group must be selected.",
                    Duration = 4000
                });
                return;
            }

            if (IsAdd)
            {
                bool exists = await ItemService.ItemIdExistsAsync(Item.ItemID);
                if (exists)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary  = "Duplicate",
                        Detail   = $"Item Code '{Item.ItemID}' already exists.",
                        Duration = 5000
                    });
                    return;
                }
            }

            Item.ForgingWeight ??= 0;
            Item.FinishedWeight ??= 0;
            Item.WasteVisible ??= 0;

            IsSaving = true;
            try
            {
                bool success = await ItemService.SaveItemAsync(Item, IsAdd);

                if (success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary  = "Saved",
                        Detail   = $"Item '{Item.ItemID}' saved successfully.",
                        Duration = 3000
                    });
                    NavigationManager.NavigateTo(!string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/company/items");
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary  = "Error",
                        Detail   = "Save failed. Please check the data and try again.",
                        Duration = 5000
                    });
                }
            }
            finally
            {
                IsSaving = false;
            }
        }

        public void Cancel() => NavigationManager.NavigateTo(!string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/company/items");
    }
}
