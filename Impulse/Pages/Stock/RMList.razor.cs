using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Impulse.Services;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Pages.Stock
{
    public partial class RMList : ComponentBase
    {
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private INotificationService NotificationService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        /// <summary>Legacy frmRM.ShowMe(p_bSampling) — /rmlist vs /rmlist/true</summary>
        [Parameter] public bool Sampling { get; set; }

        private bool isLoading = true;
        private bool showInactive = false;   // legacy mnuShowAll (checked = show all incl. inactive)
        private bool showPicture = false;
        private string searchText = string.Empty;

        // Dropdown Lists
        private List<RMGroupViewModel> groups = new();
        private List<UsedInItemModel> usedInItems = new();
        private List<RMLookupModel> rmLookup = new();

        // Selected Filter Values
        private RMGroupViewModel? selectedGroup = null;
        private UsedInItemModel? selectedUsedIn = null;
        private RMLookupModel? selectedLookup = null;

        // Items List & Paging
        private List<RMListItemModel> allItems = new();
        private List<RMListItemModel> pagedItems = new();
        private RMListItemModel? highlightedItem = null;

        private string rmPicBase64 = string.Empty;

        // Totals (legacy txtActiveArticles / txtTotalValue)
        private double totalValue = 0;

        // Paging Parameters
        private int currentPage = 1;
        private int pageSize = 50;
        private int totalPages = 1;
        private int totalCount = 0;

        private int SamplingFlag => Sampling ? 1 : 0;

        protected override async Task OnInitializedAsync()
        {
            await LoadFilters();
            await LoadData();
        }

        private async Task LoadFilters()
        {
            try
            {
                groups = await DbHelper.GetListAsync<RMGroupViewModel>("ID, Description", "RMGroups", "ORDER BY Description");
                usedInItems = await DbHelper.GetListAsync<UsedInItemModel>("DISTINCT ItemID", "ItemsRMComp", "ORDER BY ItemID");
                rmLookup = await DbHelper.GetListAsync<RMLookupModel>(
                    "RMID, RMID1, RMName, RMSize, SizeUnit",
                    "RM",
                    $"WHERE Sampling = {SamplingFlag} ORDER BY RMID1");
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
                // Mirrors legacy FillItems() condition building
                string condition = $"WHERE Sampling = {SamplingFlag}";

                if (selectedGroup != null)
                {
                    condition += $" AND GroupID = {selectedGroup.ID}";
                }

                if (!showInactive)
                {
                    condition += " AND InActive = 0";
                }

                if (selectedUsedIn != null && !string.IsNullOrEmpty(selectedUsedIn.ItemID))
                {
                    string itemId = selectedUsedIn.ItemID.Replace("'", "''");
                    condition += $" AND VRM.RMID IN (SELECT RMID FROM ItemsRMComp WHERE ItemID IN ('{itemId}'))";
                }

                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    string s = searchText.Replace("'", "''");
                    condition += $" AND (RMID1 LIKE '%{s}%' OR RMName LIKE '%{s}%')";
                }

                allItems = await DbHelper.GetListAsync<RMListItemModel>(
                    "RMID, RMID1, RMName, RMSize, SizeUnit, Unit, Rate, QtyInStock, OpenPOsQty, GroupName, GroupID, GroupColor, ReorderPoint, MinLevel, MaxLevel, InActive, Weight_Length",
                    "VRMWithRMOpenPOs VRM",
                    condition + " ORDER BY RMID1"
                );

                totalCount = allItems.Count;
                totalValue = allItems.Sum(i => i.QtyInStock * (i.Rate ?? 0));
                totalPages = (int)Math.Ceiling((double)totalCount / pageSize);
                currentPage = 1;

                UpdatePagedList();
                ResetSelection();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Raw Materials", ex.Message);
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
            rmPicBase64 = string.Empty;
        }

        private async Task OpenToolbarOptions(Microsoft.AspNetCore.Components.Web.MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15, highlightedItem);
        }

        private async Task OnGroupChanged(RMGroupViewModel? group)
        {
            selectedGroup = group;
            await LoadData();
        }

        private async Task OnUsedInChanged(UsedInItemModel? usedIn)
        {
            selectedUsedIn = usedIn;
            await LoadData();
        }

        // Legacy GoToSelected(): locate the picked RM in the grid, jump to its page and highlight it
        private void OnLookupChanged(RMLookupModel? lookup)
        {
            selectedLookup = lookup;
            if (lookup == null) return;

            int index = allItems.FindIndex(i => i.RMID == lookup.RMID);
            if (index < 0)
            {
                NotificationService.ShowInformation("Not In View", $"{lookup.RMID1} is not in the current filtered list.");
                return;
            }

            currentPage = (index / pageSize) + 1;
            UpdatePagedList();
            var item = allItems[index];
            highlightedItem = item;
            if (showPicture)
            {
                _ = LoadRMPicture(item.RMID);
            }
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
                rmPicBase64 = string.Empty;
            }
            else if (highlightedItem != null)
            {
                _ = LoadRMPicture(highlightedItem.RMID);
            }
        }

        private async Task SelectItem(RMListItemModel item)
        {
            highlightedItem = item;
            if (showPicture)
            {
                await LoadRMPicture(item.RMID);
            }
        }

        // Legacy LV_ItemClick(): loads Pic from RM table
        private async Task LoadRMPicture(int rmId)
        {
            rmPicBase64 = string.Empty;
            try
            {
                var picBytes = await DbHelper.GetSingleValueAsync<byte[]>("Pic", "RM", $"WHERE RMID = {rmId}");
                if (picBytes != null && picBytes.Length > 0)
                {
                    rmPicBase64 = Convert.ToBase64String(picBytes);
                }
            }
            catch
            {
                rmPicBase64 = string.Empty;
            }
        }

        private async Task RefreshList(ItemClickEventArgs? e = null)
        {
            await LoadFilters();
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
        private async Task<IEnumerable<RMGroupViewModel>> SearchGroups(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return groups;
            return await Task.FromResult(groups.Where(g => g.Description.Contains(text, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task<IEnumerable<UsedInItemModel>> SearchUsedInItems(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return usedInItems;
            return await Task.FromResult(usedInItems.Where(u => u.ItemID.Contains(text, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task<IEnumerable<RMLookupModel>> SearchRMLookup(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return rmLookup;
            return await Task.FromResult(rmLookup.Where(r =>
                r.RMID1.Contains(text, StringComparison.OrdinalIgnoreCase) ||
                (r.RMName?.Contains(text, StringComparison.OrdinalIgnoreCase) ?? false)));
        }

        // ── Helper to resolve Item Context from BlazorContextMenu ──
        private RMListItemModel? GetItemContext(ItemClickEventArgs? e)
        {
            if (e?.Data is RMListItemModel item)
            {
                highlightedItem = item;
                if (showPicture)
                {
                    _ = LoadRMPicture(item.RMID);
                }
            }
            return highlightedItem;
        }

        // ── Helper to build Report Selection Formulas (legacy mnuPrint_Click pattern) ──
        private string GetSelectionFormula()
        {
            string formula = "{VRM.InActive}=False";
            if (showInactive)
            {
                formula = "True";
            }
            if (selectedGroup != null)
            {
                formula += $" AND {{VRM.GroupID}}={selectedGroup.ID}";
            }
            return formula;
        }

        // ── Row Context Menu Actions ──

        private void EditRM(RMListItemModel item)
        {
            if (item.InActive)
            {
                // Legacy CheckMenu disables Edit for inactive materials
                NotificationService.ShowWarning("Inactive Material", "This raw material is inactive. Make it active first to edit.");
                return;
            }
            NavigationManager.NavigateTo($"/newrm/{item.RMID}");
        }

        private void EditRM(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No material selected", "Please select a raw material to edit.");
                return;
            }
            EditRM(item);
        }

        // Legacy MakeAccountInActive / MakeAccountActive
        private async Task ToggleActive(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No material selected", "Please select a raw material first.");
                return;
            }
            try
            {
                int newStatus = item.InActive ? 0 : 1;
                await DbHelper.ExecuteAsync("UPDATE RM SET InActive = @Status WHERE RMID = @RMID", new { Status = newStatus, RMID = item.RMID });

                NotificationService.ShowSuccess("Success", $"{item.RMID1} is now {(newStatus == 1 ? "inactive" : "active")}.");
                await LoadData();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error updating status", ex.Message);
            }
        }

        // Legacy mnuDeleteRawMaterial_Click
        private async Task DeleteRM(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No material selected", "Please select a raw material to delete.");
                return;
            }

            bool confirm = await NotificationService.ShowQuestionNotification("Delete Raw Material", "Are you sure to delete?");
            if (!confirm) return;

            try
            {
                // First check if it's assigned to any Item
                int count = await DbHelper.getSingleIntValueasync("COUNT(RMID)", "ItemsRMComp", $"WHERE RMID = {item.RMID}");
                if (count > 0)
                {
                    NotificationService.ShowInformation("Cannot Delete", "This raw material is assigned to Item(s).");
                    return;
                }

                // VendOrderDetail.RMID stores the RM code (RMID1), per legacy check on LV.SelectedItem.Text
                count = await DbHelper.getSingleIntValueasync("COUNT(RMID)", "VendOrderDetail", $"WHERE RMID = '{item.RMID1.Replace("'", "''")}'");
                if (count > 0)
                {
                    NotificationService.ShowInformation("Cannot Delete", "This raw material appears in PO(s).");
                    return;
                }

                await DbHelper.ExecuteAsync("DELETE FROM RM WHERE RMID = @RMID", new { RMID = item.RMID });
                NotificationService.ShowSuccess("Deleted", $"{item.RMID1} deleted successfully.");
                await LoadData();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error deleting", ex.Message);
            }
        }

        private void AddNewRMCopy(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No material selected", "Please select a raw material to copy.");
                return;
            }
            // Legacy frmNewRM.AddNew(bSampling, True, lRMID_For_Copy)
            NavigationManager.NavigateTo($"/newrm/copy/{item.RMID}");
        }

        private void ViewStockHistory(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No material selected", "Please select a raw material to view stock history.");
                return;
            }
            // Legacy opens frmStockLedger — not yet migrated
            NotificationService.ShowInformation("Stock History", $"Stock ledger for {item.RMID1} is not yet available.");
        }

        private void NewPurchaseOrder(ItemClickEventArgs? e = null)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No material selected", "Please select a raw material first.");
                return;
            }
            // Legacy opens frmRMPurchOrd — not yet migrated
            NotificationService.ShowInformation("Purchase Order", $"Purchase order for {item.RMID1} is not yet available.");
        }

        // ── Options Menu Actions ──

        private void NewRM(ItemClickEventArgs? e = null)
        {
            NavigationManager.NavigateTo("/newrm");
        }

        private void NewCategory(ItemClickEventArgs? e = null)
        {
            NavigationManager.NavigateTo("/stock/rmgroups");
        }

        private void ShowAll(ItemClickEventArgs? e = null)
        {
            selectedGroup = null;
            selectedUsedIn = null;
            searchText = string.Empty;
            showInactive = true;
            _ = LoadData();
        }

        // ── Row Reports (legacy PrintRMReports) ──

        private async Task PrintRowReport(string reportName, ItemClickEventArgs? e)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No material selected", "Please select a raw material to print.");
                return;
            }
            try
            {
                var request = new ReportRequest
                {
                    ReportName = reportName,
                    SelectionFormula = $"{{VRM.RMID}}={item.RMID}"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintSpecification(ItemClickEventArgs? e = null) => await PrintRowReport("RM_Specification.rpt", e);
        private async Task PrintDetails(ItemClickEventArgs? e = null) => await PrintRowReport("RM_Details.rpt", e);
        private async Task PrintLabel(ItemClickEventArgs? e = null) => await PrintRowReport("RM_Lable.rpt", e);
        private async Task PrintRLabel(ItemClickEventArgs? e = null) => await PrintRowReport("RM_R_Lable.rpt", e);
        private async Task PrintBarcode(ItemClickEventArgs? e = null) => await PrintRowReport("RM_Barcode.rpt", e);

        // ── List Reports (legacy mnuPrint* handlers) ──

        private async Task PrintReport(string reportName, string selectionFormula)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = reportName,
                    SelectionFormula = selectionFormula
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        // Legacy mnuPrint_Click → ShowRMListrpt (mod_Rpts.bas)
        private async Task PrintRMList(ItemClickEventArgs? e = null) => await PrintReport("RMList.rpt", GetSelectionFormula());

        // Legacy mnuPrintNonZero_Click
        private async Task PrintNonZero(ItemClickEventArgs? e = null)
        {
            string sel = "{VRM.QtyInStock}>0";
            if (selectedGroup != null)
            {
                sel += $" AND {{Material.GroupID}}={selectedGroup.ID}";
            }
            await PrintReport("RMList_NonZero.rpt", sel);
        }

        // Legacy mnuBelowMinLevel_Click
        private async Task PrintBelowMin(ItemClickEventArgs? e = null)
            => await PrintReport("RMList_Min.rpt", "{VRM.InActive}=FALSE AND {VRM.QtyInStock}<{VRM.MinLevel}");

        // Legacy mnuAboveMaxLevel_Click
        private async Task PrintAboveMax(ItemClickEventArgs? e = null)
            => await PrintReport("RMList_Max.rpt", "{VRM.InActive}=FALSE AND {VRM.QtyInStock}>{VRM.MaxLevel} AND {VRM.MaxLevel}>0");

        // Legacy mnuPrintRMListWithValue_Click
        private async Task PrintListWithValue(ItemClickEventArgs? e = null) => await PrintReport("RMList_Value.rpt", GetSelectionFormula());

        // Legacy mnuPrintRMListWithUsedIn_Click / CrossNo / CrossReference pass the current grid SQL (strSQL_Reports);
        // the report API only supports selection formulas, so the current filter formula is used instead.
        private async Task PrintListWithUsedIn(ItemClickEventArgs? e = null) => await PrintReport("RMList_UsedIn.rpt", GetSelectionFormula());
        private async Task PrintListCrossNo(ItemClickEventArgs? e = null) => await PrintReport("RMList_CrossNo.rpt", GetSelectionFormula());
        private async Task PrintCrossReference(ItemClickEventArgs? e = null) => await PrintReport("RMList_CrossReference.rpt", GetSelectionFormula());

        // Legacy mnuPrintRMRackwise_Click
        private async Task PrintRackwise(ItemClickEventArgs? e = null)
        {
            string sel = "{VRM.QtyInStock}>0";
            if (selectedGroup != null)
            {
                sel += $" AND {{Material.GroupID}}={selectedGroup.ID}";
            }
            await PrintReport("RMListRackwise.rpt", sel);
        }

        // Legacy mnuPrintRateHistory_Click
        private async Task PrintRateHistory(ItemClickEventArgs? e = null)
        {
            string sel = selectedGroup != null ? $"{{Material.GroupID}}={selectedGroup.ID}" : "";
            await PrintReport("RMRateHistory.rpt", sel);
        }
    }

    // ── Models ──

    public class RMListItemModel
    {
        public int RMID { get; set; }
        public string RMID1 { get; set; } = string.Empty;
        public string? RMName { get; set; }
        public float? RMSize { get; set; }
        public string? SizeUnit { get; set; }
        public string? Unit { get; set; }
        public double? Rate { get; set; }
        public float QtyInStock { get; set; }
        public float? OpenPOsQty { get; set; }
        public string? GroupName { get; set; }
        public int? GroupID { get; set; }
        public int? GroupColor { get; set; }
        public int? ReorderPoint { get; set; }
        public float? MinLevel { get; set; }
        public float? MaxLevel { get; set; }
        public bool InActive { get; set; }
        public string? Weight_Length { get; set; }

        /// <summary>Legacy: inactive rows grey (&amp;H707070), otherwise OLE COLORREF of the group.</summary>
        public string NameColorCss
        {
            get
            {
                if (InActive) return "#707070";
                int c = GroupColor ?? 0;
                int r = c & 0xFF;
                int g = (c >> 8) & 0xFF;
                int b = (c >> 16) & 0xFF;
                return $"rgb({r},{g},{b})";
            }
        }
    }

    public class UsedInItemModel
    {
        public string ItemID { get; set; } = string.Empty;
    }

    public class RMLookupModel
    {
        public int RMID { get; set; }
        public string RMID1 { get; set; } = string.Empty;
        public string? RMName { get; set; }
        public float? RMSize { get; set; }
        public string? SizeUnit { get; set; }

        /// <summary>Legacy AddHeads format: {RMID1} RMName {Size SizeUnit}</summary>
        public string DisplayText =>
            $"{{{RMID1}}} {RMName} {{{(RMSize.HasValue ? RMSize.Value.ToString("0.##") : "")} {SizeUnit}}}";
    }
}
