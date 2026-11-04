using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Production.ProductionItemList
{
    public partial class ProductionItemList : ComponentBase
    {
        [Inject] private IProductionItemListService ItemListService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;

        // ── State ────────────────────────────────────────────────────────────
        public ProductionItemListFilters Filters { get; set; } = new();
        public List<ProductionItemListItem> Items { get; set; } = new();
        public ProductionItemListLookups Lookups { get; set; } = new();
        public bool IsLoading { get; set; } = false;

        // Cascade lookup state
        public List<LookupItemString> Countries { get; set; } = new();
        public List<LookupItemString> Orders { get; set; } = new();

        // Selected typeahead bindings
        public LookupItemString? SelectedCustomer { get; set; }
        public LookupItemString? SelectedOrder { get; set; }
        public LookupItemInt? SelectedProcess { get; set; }
        public LookupItemInt? SelectedMaker { get; set; }
        public LookupItemString? SelectedItem { get; set; }

        // Right-click target
        private ProductionItemListItem? ContextItem { get; set; }

        // Date range disabled flag
        public bool IsDateRangeDisabled => Filters.DateRangeMode != 4;

        // ── Lifecycle ────────────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            try
            {
                Lookups = await ItemListService.GetInitialLookupsAsync();
                Countries = await ItemListService.GetCountriesByCustomerAsync("0");
                Orders = await ItemListService.GetOrdersByCustomerAsync("0");

                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Load Error", ex.Message);
            }
        }

        // ── Data Load ────────────────────────────────────────────────────────
        public async Task LoadDataAsync()
        {
            IsLoading = true;
            StateHasChanged();

            try
            {
                // Sync selected typeahead values into filter model before querying
                Filters.CustCode = SelectedCustomer?.Id ?? "0";
                Filters.OrderNo = SelectedOrder?.Id ?? "0";
                Filters.ProcessId = SelectedProcess?.Id.ToString() ?? "0";
                Filters.MakerId = SelectedMaker?.Id.ToString() ?? "0";
                Filters.ItemCode = SelectedItem?.Id ?? "0";

                Items = await ItemListService.GetListAsync(Filters);
                for (int i = 0; i < Items.Count; i++)
                {
                    Items[i].SNo = i + 1;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error Loading Data", ex.Message);
                Items = new();
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        // ── Filter Cascade ───────────────────────────────────────────────────
        public async Task OnCustomerChangedAsync()
        {
            var custCode = SelectedCustomer?.Id ?? "0";
            Filters.CustCode = custCode;

            // Reload cascading dropdowns in parallel
            var countriesTask = ItemListService.GetCountriesByCustomerAsync(custCode);
            var ordersTask = ItemListService.GetOrdersByCustomerAsync(custCode);
            await Task.WhenAll(countriesTask, ordersTask);

            Countries = countriesTask.Result;
            Orders = ordersTask.Result;

            // Reset selections that depend on customer
            Filters.Country = "0";
            SelectedOrder = null;
            Filters.OrderNo = "0";

            StateHasChanged();
        }

        public void OnDateRangeModeChanged()
        {
            if (Filters.DateRangeMode != 4)
            {
                Filters.DtFrom = DateTime.Today;
                Filters.DtTo = DateTime.Today;
            }
        }

        public void OpenOptionsMenu()
        {
            // The ContextMenuTrigger handles this automatically
        }

        // ── Row Selection ───────────────────────────────────────────────────
        public bool AreAllSelected => Items.Count > 0 && Items.All(x => x.IsSelected);
        public int SelectedCount => Items.Count(x => x.IsSelected);

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            bool selectAll = (bool)(e.Value ?? false);
            foreach (var item in Items)
            {
                item.IsSelected = selectAll;
            }
            StateHasChanged();
        }

        public void SelectAll(ItemClickEventArgs? args = null)
        {
            foreach (var item in Items)
            {
                item.IsSelected = true;
            }
            StateHasChanged();
        }

        public void ClearAll(ItemClickEventArgs? args = null)
        {
            foreach (var item in Items)
            {
                item.IsSelected = false;
            }
            StateHasChanged();
        }

        // ── Typeahead Search Methods ─────────────────────────────────────────
        public Task<IEnumerable<LookupItemString>> SearchCustomers(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? Lookups.Customers
                : Lookups.Customers.Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult(results.AsEnumerable());
        }

        public Task<IEnumerable<LookupItemString>> SearchOrders(string searchText)
        {
            var source = Orders.Count > 0 ? Orders : Lookups.Customers; // fallback if not loaded
            var results = string.IsNullOrWhiteSpace(searchText)
                ? (Orders.Count > 0 ? Orders : new List<LookupItemString>())
                : Orders.Where(o => o.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult(results.AsEnumerable());
        }

        public Task<IEnumerable<LookupItemInt>> SearchProcesses(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? Lookups.Processes
                : Lookups.Processes.Where(p => p.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult(results.AsEnumerable());
        }

        public Task<IEnumerable<LookupItemInt>> SearchMakers(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? Lookups.Makers
                : Lookups.Makers.Where(m => m.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult(results.AsEnumerable());
        }

        public Task<IEnumerable<LookupItemString>> SearchItems(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? Lookups.Items
                : Lookups.Items.Where(i => i.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult(results.AsEnumerable());
        }

        // ── Inline Remarks Editing ───────────────────────────────────────────
        public void StartRemarksEdit(ProductionItemListItem item)
        {
            // Cancel any other edits in progress
            foreach (var other in Items.Where(i => i.IsEditingRemarks && i != item))
                CancelRemarksEdit(other);

            item.EditingRemarksValue = item.DisplayRemarks;
            item.IsEditingRemarks = true;
        }

        public void CancelRemarksEdit(ProductionItemListItem item)
        {
            item.IsEditingRemarks = false;
        }

        public async Task SaveRemarksOnBlur(ProductionItemListItem item)
        {
            if (!item.IsEditingRemarks) return;
            await CommitRemarksAsync(item);
        }

        public async Task OnRemarksKeyDown(KeyboardEventArgs e, ProductionItemListItem item)
        {
            if (e.Key == "Enter")
            {
                await CommitRemarksAsync(item);
            }
            else if (e.Key == "Escape")
            {
                CancelRemarksEdit(item);
            }
        }

        private async Task CommitRemarksAsync(ProductionItemListItem item)
        {
            item.IsEditingRemarks = false;

            // If value did not change, no need to execute DB update
            if (item.EditingRemarksValue == item.DisplayRemarks) return;

            try
            {
                bool saved = await ItemListService.SaveRemarksAsync(
                    item.LotNo, item.EntryType, item.OrderNo, item.ItemCode, item.EditingRemarksValue);

                if (saved)
                {
                    item.POIRemarks = item.EditingRemarksValue;
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", "Remarks updated.");
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Warning, "Remarks", "No rows updated.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Save Error", ex.Message);
            }
        }

        // ── Print Actions ────────────────────────────────────────────────────
        // The print reports use SQLQueryString against the Crystal report.
        // We pass the active filter as selection formula to the Blazor report API.

        public async Task PrintListAsync(ItemClickEventArgs args)
        {
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "ProductionItemList.rpt",
                SelectionFormula = BuildCrystalSelectionFormula()
            });
        }

        public async Task PrintSelectedAsync(ItemClickEventArgs args)
        {
            var selected = Items.Where(i => i.IsSelected).ToList();
            if (selected.Count == 0)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Print (Selected)", "Please select at least one row to print.");
                return;
            }

            var formula = BuildSelectedSelectionFormula(selected);
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "ProductionItemList.rpt",
                SelectionFormula = formula
            });
        }

        // ── Right-Click: Print PTC ────────────────────────────────────────────
        public async Task PrintPTCFromRow(ItemClickEventArgs args)
        {
            var item = args.Data as ProductionItemListItem;
            if (item == null || string.IsNullOrWhiteSpace(item.LotNo))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Print PTC", "No lot selected.");
                return;
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQel.rpt",
                Parameters = new Dictionary<string, object>
                {
                    { "@LotNo", item.LotNo }
                }
            });
        }

        // ── Helpers ──────────────────────────────────────────────────────────
        /// <summary>
        /// Builds a Crystal Reports selection formula that mirrors the current active filters.
        /// </summary>
        private string BuildCrystalSelectionFormula()
        {
            var parts = new List<string> { "{VRunningLots_POI.Qty} > 0" };

            if (!string.IsNullOrWhiteSpace(Filters.LotNo))
            {
                parts.Clear();
                parts.Add($"{{VRunningLots_POI.LotNo}} = '{Filters.LotNo}'");
                return string.Join(" AND ", parts);
            }

            if (Filters.DateRangeMode != 0)
            {
                DateTime dtFrom = Filters.DtFrom;
                DateTime dtTo = Filters.DtTo;
                if (Filters.DateRangeMode == 1) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(15); }
                else if (Filters.DateRangeMode == 2) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(30); }
                else if (Filters.DateRangeMode == 3) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(45); }
                parts.Add($"{{VRunningLots_POI.DT}} in Date({dtFrom.Year}, {dtFrom.Month}, {dtFrom.Day}) to Date({dtTo.Year}, {dtTo.Month}, {dtTo.Day})");
            }

            if (!string.IsNullOrEmpty(Filters.CustCode) && Filters.CustCode != "0")
                parts.Add($"{{VRunningLots_POI.CustCode}} = '{Filters.CustCode}'");

            if (!string.IsNullOrEmpty(Filters.Country) && Filters.Country != "0")
                parts.Add($"{{VRunningLots_POI.Country}} = '{Filters.Country}'");

            if (!string.IsNullOrEmpty(Filters.OrderNo) && Filters.OrderNo != "0")
                parts.Add($"{{VRunningLots_POI.OrderNo}} = '{Filters.OrderNo}'");

            if (!string.IsNullOrEmpty(Filters.ItemGroupId) && Filters.ItemGroupId != "0")
                parts.Add($"{{VRunningLots_POI.GroupID}} = {Filters.ItemGroupId}");

            if (!string.IsNullOrEmpty(Filters.Hub) && Filters.Hub != "0")
                parts.Add($"{{VRunningLots_POI.Hub_Name}} = '{Filters.Hub}'");

            if (!Filters.ShowRepairLots) parts.Add("{VRunningLots_POI.ReWorkLot} = 0");
            if (!Filters.ShowRegularLots) parts.Add("{VRunningLots_POI.ReWorkLot} = 1");

            if (Filters.ShowLotsOnly && Filters.ShowNotYetStarted)
            {
                parts.Add("(({VRunningLots_POI.EntryType} <> 0 AND (IsNull({VRunningLots_POI.DispatchQty}) OR {VRunningLots_POI.DispatchQty} < {VRunningLots_POI.OrderQty})) OR {VRunningLots_POI.EntryType} <> 1)");
            }
            else
            {
                if (!Filters.ShowLotsOnly)
                    parts.Add("({VRunningLots_POI.EntryType} <> 0 AND (IsNull({VRunningLots_POI.DispatchQty}) OR {VRunningLots_POI.DispatchQty} < {VRunningLots_POI.OrderQty}))");
                if (!Filters.ShowNotYetStarted)
                    parts.Add("{VRunningLots_POI.EntryType} <> 1");
            }

            return string.Join(" AND ", parts);
        }

        /// <summary>
        /// Builds a Crystal Reports selection formula restricted exclusively to the selected/marked rows.
        /// </summary>
        private string BuildSelectedSelectionFormula(List<ProductionItemListItem> selected)
        {
            var conditions = new List<string>();

            // 1. Lots (EntryType = 0) with valid LotNo
            var lotNos = selected
                .Where(x => !string.IsNullOrWhiteSpace(x.LotNo))
                .Select(x => x.LotNo.Trim())
                .Distinct()
                .ToList();

            if (lotNos.Count > 0)
            {
                var lotList = string.Join(", ", lotNos.Select(l => $"'{l.Replace("'", "''")}'"));
                conditions.Add($"{{VRunningLots_POI.LotNo}} in [{lotList}]");
            }

            // 2. Not-yet-started orders (EntryType = 1) identified by OrderNo + ItemCode
            var nonLots = selected
                .Where(x => string.IsNullOrWhiteSpace(x.LotNo) && !string.IsNullOrWhiteSpace(x.OrderNo) && !string.IsNullOrWhiteSpace(x.ItemCode))
                .ToList();

            if (nonLots.Count > 0)
            {
                var orderClauses = nonLots.Select(x =>
                    $"({{VRunningLots_POI.EntryType}} = 1 AND {{VRunningLots_POI.OrderNo}} = '{x.OrderNo.Replace("'", "''")}' AND {{VRunningLots_POI.ItemCode}} = '{x.ItemCode.Replace("'", "''")}')"
                );
                conditions.Add($"({string.Join(" OR ", orderClauses)})");
            }

            return conditions.Count > 0 ? string.Join(" OR ", conditions) : "1=0";
        }
    }
}
