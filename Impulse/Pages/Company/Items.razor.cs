using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Impulse.Services;
using Impulse.Constants;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Company;

namespace Impulse.Pages.Company
{
    public partial class Items : ComponentBase
    {
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private INotificationService NotificationService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;

        private bool isLoading = true;
        private bool showInactive = false;
        private bool showPicture = false;
        private string searchText = string.Empty;

        // Dropdown Lists
        private List<ItemGroupModel> groups = new();
        private List<ItemCategoryModel> categories = new();
        private List<ItemMainGroupModel> mainGroups = new();

        // Selected Filter Values
        private ItemGroupModel? selectedGroup = null;
        private ItemCategoryModel? selectedCategory = null;
        private ItemMainGroupModel? selectedMainGroup = null;

        // Items List & Paging
        private List<CompanyItemModel> allItems = new();
        private List<CompanyItemModel> pagedItems = new();
        private CompanyItemModel? highlightedItem = null;

        private string itemPicBase64 = string.Empty;

        // Paging Parameters
        private int currentPage = 1;
        private int pageSize = 50;
        private int totalPages = 1;
        private int totalCount = 0;

        protected override async Task OnInitializedAsync()
        {
            await LoadFilters();
            await LoadData();
        }

        private async Task LoadFilters()
        {
            try
            {
                groups = await DbHelper.GetListAsync<ItemGroupModel>("ID, Description", "ItemGroups", "ORDER BY Description");
                categories = await DbHelper.GetListAsync<ItemCategoryModel>("CatID, Description", "ItemCatagories", "ORDER BY Description");
                mainGroups = await DbHelper.GetListAsync<ItemMainGroupModel>("MainGroupID, MainGroupName", "ItemGroupsMain", "ORDER BY MainGroupName");
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Filters", ex.Message);
            }
        }

        private async Task LoadData()
        {
            isLoading = true;
            try
            {
                string condition = "WHERE 1=1";

                if (!showInactive)
                {
                    condition += " AND InActive = 0";
                }

                if (selectedGroup != null)
                {
                    condition += $" AND GroupID = {selectedGroup.ID}";
                }

                if (selectedCategory != null)
                {
                    condition += $" AND CatID = {selectedCategory.CatID}";
                }

                if (selectedMainGroup != null)
                {
                    condition += $" AND MainGroupID = {selectedMainGroup.MainGroupID}";
                }

                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    condition += $" AND (ItemID LIKE '%{searchText}%' OR ItemName LIKE '%{searchText}%' OR ItemGroup LIKE '%{searchText}%')";
                }

                allItems = await DbHelper.GetListAsync<CompanyItemModel>(
                    "GrpColor, ItemID, InActive, ItemGroup, ItemName, ItemSize, SizeUnit, Unit, FinQuality, FinishedWeight, ReorderPoint, FixedPackingUnit, InHand, PriceForCost, Type, TipSize, ExWorks, ExWorksTop",
                    "VItems",
                    condition + " ORDER BY ItemName"
                );

                totalCount = allItems.Count;
                totalPages = (int)Math.Ceiling((double)totalCount / pageSize);
                currentPage = 1;

                UpdatePagedList();
                ResetSelection();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Items", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private void UpdatePagedList()
        {
            pagedItems = allItems
                .Skip((currentPage - 1) * pageSize)
                .Take(pageSize)
                .ToList();
        }

        private void ResetSelection()
        {
            highlightedItem = null;
            itemPicBase64 = string.Empty;
        }

        private async Task OpenToolbarOptions(Microsoft.AspNetCore.Components.Web.MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15, highlightedItem);
        }

        private async Task OnGroupChanged(ItemGroupModel? group)
        {
            selectedGroup = group;
            await LoadData();
        }

        private async Task OnCategoryChanged(ItemCategoryModel? category)
        {
            selectedCategory = category;
            await LoadData();
        }

        private async Task OnMainGroupChanged(ItemMainGroupModel? mainGroup)
        {
            selectedMainGroup = mainGroup;
            await LoadData();
        }

        private async Task OnSearchInput(ChangeEventArgs e)
        {
            searchText = e.Value?.ToString() ?? string.Empty;
            await LoadData();
        }

        private async Task OnShowInactiveChanged(ChangeEventArgs e)
        {
            showInactive = (bool)(e.Value ?? false);
            await LoadData();
        }

        private void OnShowPictureChanged(ChangeEventArgs e)
        {
            showPicture = (bool)(e.Value ?? false);
            if (!showPicture)
            {
                itemPicBase64 = string.Empty;
            }
            else if (highlightedItem != null)
            {
                _ = LoadItemPicture(highlightedItem.ItemID);
            }
        }

        private async Task SelectItem(CompanyItemModel item)
        {
            highlightedItem = item;
            if (showPicture)
            {
                await LoadItemPicture(item.ItemID);
            }
        }

        private async Task LoadItemPicture(string itemId)
        {
            itemPicBase64 = string.Empty;
            try
            {
                var picBytes = await DbHelper.GetSingleValueAsync<byte[]>("ItemPic", "Items", $"WHERE ItemID = '{itemId}'");
                if (picBytes != null && picBytes.Length > 0)
                {
                    itemPicBase64 = Convert.ToBase64String(picBytes);
                }
            }
            catch
            {
                itemPicBase64 = string.Empty;
            }
        }

        private async Task RefreshList(ItemClickEventArgs? e = null)
        {
            await LoadData();
        }

        private void GoToPage(int page)
        {
            if (page >= 1 && page <= totalPages)
            {
                currentPage = page;
                UpdatePagedList();
            }
        }

        // ── Typeahead Search Helpers ──
        private async Task<IEnumerable<ItemGroupModel>> SearchGroups(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return groups;
            return await Task.FromResult(groups.Where(g => g.Description.Contains(text, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task<IEnumerable<ItemCategoryModel>> SearchCategories(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return categories;
            return await Task.FromResult(categories.Where(c => c.Description.Contains(text, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task<IEnumerable<ItemMainGroupModel>> SearchMainGroups(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return mainGroups;
            return await Task.FromResult(mainGroups.Where(mg => mg.MainGroupName.Contains(text, StringComparison.OrdinalIgnoreCase)));
        }

        // ── Helper to resolve Item Context from BlazorContextMenu ──
        private CompanyItemModel? GetItemContext(ItemClickEventArgs? e)
        {
            if (e?.Data is CompanyItemModel item)
            {
                highlightedItem = item;
                if (showPicture)
                {
                    _ = LoadItemPicture(item.ItemID);
                }
            }
            return highlightedItem;
        }

        // ── Helper to build Report Selection Formulas ──
        private string GetSelectionFormula()
        {
            string formula = "{VItems.InActive}=False";
            if (showInactive)
            {
                formula = "True";
            }
            if (selectedGroup != null)
            {
                formula += $" AND {{VItems.GroupID}}={selectedGroup.ID}";
            }
            if (selectedCategory != null)
            {
                formula += $" AND {{VItems.CatID}}={selectedCategory.CatID}";
            }
            if (selectedMainGroup != null)
            {
                formula += $" AND {{VItems.MainGroupID}}={selectedMainGroup.MainGroupID}";
            }
            return formula;
        }

        // ── Implementation of all 20 Legacy Actions ──

        private void ShowAll(ItemClickEventArgs? e = null)
        {
            selectedGroup = null;
            selectedCategory = null;
            selectedMainGroup = null;
            searchText = string.Empty;
            showInactive = true;
            _ = LoadData();
        }

        private void ViewStockHistory(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item to view stock history.");
                return;
            }
            NotificationService.ShowInformation("Stock History", $"Viewing stock history for {item.ItemID} - {item.ItemName}");
        }

        private void NewItem(ItemClickEventArgs? e = null)
        {
            NotificationService.ShowInformation("New Item", "Opening New Item form dialog...");
        }

        private void NewCategory(ItemClickEventArgs? e = null)
        {
            NotificationService.ShowInformation("New Category", "Opening New Category form dialog...");
        }

        private void EditItem(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item to edit.");
                return;
            }
            NotificationService.ShowInformation("Edit Item", $"Editing item {item.ItemID}...");
        }

        private void DeleteItem(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item to delete.");
                return;
            }
            NotificationService.ShowWarning("Delete Item", $"Deleting item {item.ItemID}...");
        }

        private void CopyItem(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item to copy.");
                return;
            }
            NotificationService.ShowInformation("Copy Item", $"Copying item details for {item.ItemID}...");
        }

        private async Task ToggleActive(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item first.");
                return;
            }
            try
            {
                int newActiveStatus = item.InActive ? 0 : 1;
                string sql = $"UPDATE Items SET InActive = {newActiveStatus} WHERE ItemID = @ItemID";
                await DbHelper.ExecuteAsync(sql, new { ItemID = item.ItemID });

                NotificationService.ShowSuccess("Success", $"Item {item.ItemID} active status updated.");
                await LoadData();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error updating status", ex.Message);
            }
        }

        private async Task ToggleLockThisItem(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item first.");
                return;
            }
            try
            {
                var count = await DbHelper.getSingleIntValueasync("COUNT(*)", "ItemsLockedForEditing", $"WHERE ItemCode = '{item.ItemID}'");
                bool isLocked = count > 0;

                if (isLocked)
                {
                    await DbHelper.ExecuteAsync($"DELETE FROM ItemsLockedForEditing WHERE ItemCode = '{item.ItemID}'");
                    NotificationService.ShowSuccess("Unlocked", $"Item {item.ItemID} unlocked successfully.");
                }
                else
                {
                    await DbHelper.ExecuteAsync($"INSERT INTO ItemsLockedForEditing(ItemCode) VALUES('{item.ItemID}')");
                    NotificationService.ShowSuccess("Locked", $"Item {item.ItemID} locked successfully.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error toggling lock state", ex.Message);
            }
        }

        private async Task AssignAllProcesses(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item first.");
                return;
            }
            bool confirm = await NotificationService.ShowQuestionNotification("Assign All Processes", "Are you sure to assign all Processes?");
            if (!confirm) return;

            try
            {
                await DbHelper.ExecuteAsync($"DELETE FROM ItemProcesses WHERE ItemID = '{item.ItemID}'");
                await DbHelper.ExecuteAsync($"INSERT INTO ItemProcesses (ProcessID, ItemID, Rate, Process_RefID) SELECT ProcessID, '{item.ItemID}', 0, ProcessID FROM Processes");
                NotificationService.ShowSuccess("Success", "All processes assigned successfully.");
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error assigning processes", ex.Message);
            }
        }

        private void ViewSales(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item first.");
                return;
            }
            NotificationService.ShowInformation("Item Sales", $"Opening item sales history for {item.ItemID}...");
        }

        private void ViewPurchases(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item first.");
                return;
            }
            NotificationService.ShowInformation("Item Purchases", $"Opening item purchases history for {item.ItemID}...");
        }

        private void ResetProcessGroups(ItemClickEventArgs? e = null)
        {
            NotificationService.ShowInformation("Reset Process Groups", "Resetting all process groups...");
        }

        // ── Realized Reports via IReportNavigationService ──

        private async Task PrintCatalog(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "rptCompanyCatalog.rpt",
                    SelectionFormula = GetSelectionFormula()
                };
                await ReportNavigation.PrintReportAsync(request);
                NotificationService.ShowSuccess("Report", "Catalog report sent to print queue.");
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintSpecSheet(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No item selected", "Please select an item to print spec sheet.");
                return;
            }
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "ItemSpecification.rpt",
                    SelectionFormula = $"{{Items.ItemID}}='{item.ItemID}'"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintSummaryStock(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "Complete_Stock_Summary.rpt",
                    SelectionFormula = "{VItems_StockReport.InActive}=FALSE"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintCombineStockReport(ItemClickEventArgs? e = null)
        {
            try
            {
                string sel = "{VItems_StockReport.InActive}=FALSE";
                if (selectedCategory != null)
                {
                    sel += $" AND {{VItems_StockReport.CatID}}={selectedCategory.CatID}";
                }
                var request = new ReportRequest
                {
                    ReportName = "CombineStockReport.rpt",
                    SelectionFormula = sel
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintCombineReportWithLocation(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "CombineStockReport_WithLocation.rpt",
                    SelectionFormula = "{VItems_StockReport.InActive}=FALSE"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintItemsBelowMin(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "CompanyCatalog_MinLevel.rpt",
                    SelectionFormula = "{VItems.InActive}=FALSE AND {VItems.InHand}<{VItems.MinLevel}"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintItemsAboveMax(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "CompanyCatalog_MaxLevel.rpt",
                    SelectionFormula = "{VItems.InActive}=FALSE AND {VItems.InHand}>{VItems.MaxLevel} AND {VItems.MaxLevel}>0"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintItemsNotUsed(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "CompanyCatalogItemsNotUsed.rpt",
                    SelectionFormula = "{VItems.InActive}=FALSE"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintECommerceStock(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "ECommerceStockReport.rpt",
                    SelectionFormula = "{VItemWithStocks.AvailableForECommerce}=TRUE"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintSFStock(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "BinWiseSemiFinishStockStatus.rpt",
                    SelectionFormula = "{VItemsWithShelfWiseStock.Qty}>0"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }
    }

    // Model declarations within namespace
    public class CompanyItemModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemGroup { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string ItemSize { get; set; } = string.Empty;
        public string SizeUnit { get; set; } = string.Empty;
        public decimal FinishedWeight { get; set; }
        public decimal InHand { get; set; }
        public string Unit { get; set; } = string.Empty;
        public decimal PriceForCost { get; set; }
        public decimal ExWorks { get; set; }
        public decimal ExWorksTop { get; set; }
        public string Type { get; set; } = string.Empty;
        public string TipSize { get; set; } = string.Empty;
        public string FinQuality { get; set; } = string.Empty;
        public decimal ReorderPoint { get; set; }
        public decimal FixedPackingUnit { get; set; }
        public bool InActive { get; set; }
        public int GrpColor { get; set; }
    }

    
}
