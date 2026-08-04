using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Production.ReceivingList
{
    public partial class ReceivingList : ComponentBase
    {
        [Inject] public IMakerRcvListService RcvListService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;
        [Inject] public AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;
        [Inject] public IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;

        // ─────────────────────────────────────────────────────────────
        // State
        // ─────────────────────────────────────────────────────────────
        public MakerRcvListFilter Filter { get; set; } = new MakerRcvListFilter();
        public List<MakerRcvListItem> AllItems { get; set; } = new List<MakerRcvListItem>();
        public string ClientSearchTerm { get; set; } = string.Empty;
        public string LastReportSql { get; set; } = string.Empty;
        public bool IsLoading { get; set; } = false;
        public MakerRcvListItem? SelectedItem { get; set; } = null;

        // Manual PTC panel
        public bool ShowManualPTCPanel { get; set; } = false;
        public string ManualPTCLotNo { get; set; } = string.Empty;
        public string ManualPTCValue { get; set; } = string.Empty;

        // ─────────────────────────────────────────────────────────────
        // Multi-select typeahead bindings
        // ─────────────────────────────────────────────────────────────
        private IList<LookupItemInt> _selectedMakers = new List<LookupItemInt>();
        public IList<LookupItemInt> SelectedMakers
        {
            get => _selectedMakers;
            set
            {
                _selectedMakers = value ?? new List<LookupItemInt>();
                Filter.MakerIds = _selectedMakers.Select(m => m.Id).ToHashSet();
                _ = OnFilterChanged();
            }
        }

        private IList<LookupItemString> _selectedCategories = new List<LookupItemString>();
        public IList<LookupItemString> SelectedCategories
        {
            get => _selectedCategories;
            set
            {
                _selectedCategories = value ?? new List<LookupItemString>();
                Filter.ItemCatIds = _selectedCategories.Select(c => c.Id).ToHashSet();
                _ = OnFilterChanged();
            }
        }

        private IList<LookupItemInt> _selectedGroups = new List<LookupItemInt>();
        public IList<LookupItemInt> SelectedGroups
        {
            get => _selectedGroups;
            set
            {
                _selectedGroups = value ?? new List<LookupItemInt>();
                Filter.ItemGroupIds = _selectedGroups.Select(g => g.Id).ToHashSet();
                _ = OnFilterChanged();
            }
        }

        private IList<LookupItemInt> _selectedProcesses = new List<LookupItemInt>();
        public IList<LookupItemInt> SelectedProcesses
        {
            get => _selectedProcesses;
            set
            {
                _selectedProcesses = value ?? new List<LookupItemInt>();
                Filter.ProcessIds = _selectedProcesses.Select(p => p.Id).ToHashSet();
                _ = OnFilterChanged();
            }
        }

        private IList<LookupItemString> _selectedCustomers = new List<LookupItemString>();
        public IList<LookupItemString> SelectedCustomers
        {
            get => _selectedCustomers;
            set
            {
                _selectedCustomers = value ?? new List<LookupItemString>();
                Filter.CustomerCodes = _selectedCustomers.Select(c => c.Id).ToHashSet();
                _ = OnFilterChanged();
            }
        }

        private LookupItemString? _selectedArticle;
        public LookupItemString? SelectedArticle
        {
            get => _selectedArticle;
            set
            {
                _selectedArticle = value;
                Filter.ItemId = value?.Id ?? "0";
                _ = OnFilterChanged();
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Lookup data
        // ─────────────────────────────────────────────────────────────
        public List<LookupItemInt> Makers { get; set; } = new List<LookupItemInt>();
        public List<LookupItemString> Categories { get; set; } = new List<LookupItemString>();
        public List<LookupItemInt> Groups { get; set; } = new List<LookupItemInt>();
        public List<LookupItemInt> Processes { get; set; } = new List<LookupItemInt>();
        public List<LookupItemString> Customers { get; set; } = new List<LookupItemString>();
        public List<LookupItemString> AllItemsList { get; set; } = new List<LookupItemString>();

        // ─────────────────────────────────────────────────────────────
        // Computed helpers
        // ─────────────────────────────────────────────────────────────
        public bool HasValidLotNo => SelectedItem != null && !string.IsNullOrWhiteSpace(SelectedItem.LotNo) && SelectedItem.LotNo != "0";

        public List<MakerRcvListItem> FilteredItems
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ClientSearchTerm))
                    return AllItems;

                string term = ClientSearchTerm.Trim().ToLower();
                return AllItems.Where(i =>
                    (i.VendID1 != null && i.VendID1.ToLower().Contains(term)) ||
                    (i.VenderName != null && i.VenderName.ToLower().Contains(term)) ||
                    (i.RecieptID != null && i.RecieptID.ToLower().Contains(term)) ||
                    (i.Description != null && i.Description.ToLower().Contains(term)) ||
                    (i.FullArticle != null && i.FullArticle.ToLower().Contains(term)) ||
                    (i.LotNo != null && i.LotNo.ToLower().Contains(term)) ||
                    (i.MasterPONo != null && i.MasterPONo.ToLower().Contains(term)) ||
                    (i.OrderNo != null && i.OrderNo.ToLower().Contains(term)) ||
                    (i.BillNo != null && i.BillNo.ToLower().Contains(term))
                ).ToList();
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Lifecycle
        // ─────────────────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            await LoadLookupsAsync();
            await LoadDataAsync();
        }

        private async Task LoadLookupsAsync()
        {
            try
            {
                var makersTask = RcvListService.GetMakersAsync();
                var categoriesTask = RcvListService.GetItemCategoriesAsync();
                var groupsTask = RcvListService.GetItemGroupsAsync();
                var processesTask = RcvListService.GetProcessesAsync();
                var customersTask = RcvListService.GetCustomersAsync();
                var itemsTask = RcvListService.GetItemsAsync();

                await Task.WhenAll(makersTask, categoriesTask, groupsTask, processesTask, customersTask, itemsTask);

                Makers = await makersTask;
                Categories = await categoriesTask;
                Groups = await groupsTask;
                Processes = await processesTask;
                Customers = await customersTask;
                AllItemsList = await itemsTask;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Lookups",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task LoadDataAsync()
        {
            IsLoading = true;
            StateHasChanged();

            try
            {
                var result = await RcvListService.GetListAsync(Filter);
                AllItems = result.Items;
                LastReportSql = result.ReportSql;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Receiving List",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        public async Task OnFilterChanged()
        {
            await LoadDataAsync();
        }

        // ─────────────────────────────────────────────────────────────
        // Typeahead Search Methods
        // ─────────────────────────────────────────────────────────────
        public Task<IEnumerable<LookupItemInt>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemInt>>(Makers);
            string q = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemInt>>(Makers.Where(m => m.Name.ToLower().Contains(q)).ToList());
        }

        public Task<IEnumerable<LookupItemString>> SearchCategories(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemString>>(Categories);
            string q = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemString>>(Categories.Where(c => c.Name.ToLower().Contains(q)).ToList());
        }

        public Task<IEnumerable<LookupItemInt>> SearchGroups(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemInt>>(Groups);
            string q = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemInt>>(Groups.Where(g => g.Name.ToLower().Contains(q)).ToList());
        }

        public Task<IEnumerable<LookupItemInt>> SearchProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemInt>>(Processes);
            string q = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemInt>>(Processes.Where(p => p.Name.ToLower().Contains(q)).ToList());
        }

        public Task<IEnumerable<LookupItemString>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemString>>(Customers);
            string q = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemString>>(Customers.Where(c => c.Name.ToLower().Contains(q)).ToList());
        }

        public Task<IEnumerable<LookupItemString>> SearchItems(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemString>>(AllItemsList);
            string q = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemString>>(AllItemsList.Where(i => i.Name.ToLower().Contains(q)).ToList());
        }

        // ─────────────────────────────────────────────────────────────
        // Checkbox Handlers
        // ─────────────────────────────────────────────────────────────
        public async Task OnRepairLotsChanged()
        {
            if (Filter.RepairLots) Filter.RegularLotsOnly = false;
            await LoadDataAsync();
        }

        public async Task OnRegularLotsOnlyChanged()
        {
            if (Filter.RegularLotsOnly) Filter.RepairLots = false;
            await LoadDataAsync();
        }

        public async Task OnBilledChanged()
        {
            if (Filter.BilledOnly) Filter.NonBilledOnly = false;
            await LoadDataAsync();
        }

        public async Task OnNonBilledChanged()
        {
            if (Filter.NonBilledOnly) Filter.BilledOnly = false;
            await LoadDataAsync();
        }

        public async Task OnRefreshClicked()
        {
            await LoadDataAsync();
        }

        // ─────────────────────────────────────────────────────────────
        // Row interaction
        // ─────────────────────────────────────────────────────────────
        public void SelectRow(MakerRcvListItem item)
        {
            SelectedItem = item;
        }

        /// <summary>
        /// Returns CSS class for row coloring matching legacy VB6 color rules.
        /// </summary>
        public string GetRowClass(MakerRcvListItem item)
        {
            if (string.IsNullOrEmpty(item.CompItemCode))
                return "row-no-item";      // Red text (no item code = problem row)
            if (item.ReWorkLot)
                return "row-repair";       // Repair lot = light red tint
            if (!string.IsNullOrEmpty(item.ComplaintItemID))
                return "row-rejection";    // Complaint/rejection = yellow tint
            return string.Empty;
        }

        public async Task ShowOptionsMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is MakerRcvListItem item)
                SelectedItem = item;
        }

        private async Task<string> GetCurrentUserName()
        {
            var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
            return authState.User.Identity?.Name ?? "Admin";
        }

        // ─────────────────────────────────────────────────────────────
        // Row-Specific Context Menu Actions
        // ─────────────────────────────────────────────────────────────

        public async Task DeleteReceiving(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null) return;

            // Check for linked issuances
            int issCount = await RcvListService.CheckIssuanceExistsAsync(SelectedItem.VRD_EntryID);
            if (issCount > 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Cannot Delete",
                    Detail = $"Lot {SelectedItem.LotNo} has {issCount} linked issuance(s). Delete them first.",
                    Duration = 5000
                });
                return;
            }

            string userName = await GetCurrentUserName();
            string machineName = Environment.MachineName;

            bool success = await RcvListService.DeleteReceivingAsync(SelectedItem.VRD_EntryID, userName, machineName);
            if (success)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Deleted",
                    Detail = $"Lot {SelectedItem.LotNo} receiving deleted successfully.",
                    Duration = 4000
                });
                SelectedItem = null;
                await LoadDataAsync();
            }
        }

        public async Task CloseLot(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || !HasValidLotNo) return;

            string userName = await GetCurrentUserName();
            string machineName = Environment.MachineName;

            bool success = await RcvListService.CloseLotAsync(SelectedItem.LotNo, userName, machineName);
            if (success)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Lot Closed",
                    Detail = $"Lot {SelectedItem.LotNo} has been closed.",
                    Duration = 4000
                });
                await LoadDataAsync();
            }
            else
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Already Closed",
                    Detail = $"Lot {SelectedItem.LotNo} was already closed.",
                    Duration = 3000
                });
            }
        }

        public void OpenManualPTCPanel(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || !HasValidLotNo) return;

            ManualPTCLotNo = SelectedItem.LotNo;
            ManualPTCValue = string.Empty;
            ShowManualPTCPanel = true;
        }

        public async Task SaveManualPTCNo()
        {
            if (string.IsNullOrWhiteSpace(ManualPTCLotNo)) return;

            bool success = await RcvListService.UpdateManualPTCNoAsync(ManualPTCLotNo, ManualPTCValue);
            if (success)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Manual PTC Updated",
                    Detail = $"Lot {ManualPTCLotNo}: PTC No set to '{ManualPTCValue}'.",
                    Duration = 3000
                });
                ShowManualPTCPanel = false;
            }
        }

        public void PlaceholderAction(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Info,
                Summary = "Coming Soon",
                Detail = "This feature will be available in a future update.",
                Duration = 3000
            });
        }

        // ─────────────────────────────────────────────────────────────
        // Row-Specific Reports
        // ─────────────────────────────────────────────────────────────

        public async Task PrintSlip(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || !SelectedItem.IssuanceRefID.HasValue) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssSlipWithRcving.rpt",
                SelectionFormula = $"{{VendIssued.EntryID}} = {SelectedItem.IssuanceRefID.Value}"
            });
        }

        public async Task PrintSlipMini(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || !SelectedItem.IssuanceRefID.HasValue) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssSlipWithRcving_Mini.rpt",
                SelectionFormula = $"{{VendIssued.EntryID}} = {SelectedItem.IssuanceRefID.Value}",
                FormulaValues = new Dictionary<string, object>
                {
                    { "SubFilter", $"{{VendRcvdDetail.EntryID}} = {SelectedItem.VRD_EntryID}" }
                }
            });
        }

        public async Task PrintPTC(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasValidLotNo) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTC.rpt",
                SelectionFormula = $"{{VendRcvdDetail.EntryID}} = {SelectedItem!.VRD_EntryID}"
            });
        }

        public async Task PrintPTCMini(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasValidLotNo) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTC_Mini.rpt",
                SelectionFormula = $"{{VendRcvdDetail.EntryID}} = {SelectedItem!.VRD_EntryID}"
            });
        }

        public async Task PrintPTCWithCost(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasValidLotNo) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTC.rpt",
                SelectionFormula = $"{{VendRcvdDetail.EntryID}} = {SelectedItem!.VRD_EntryID}",
                FormulaValues = new Dictionary<string, object> { { "ShowCost", true } }
            });
        }

        public async Task PrintRepairSlip(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasValidLotNo) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "RepairSlip.rpt",
                SelectionFormula = $"{{VendRcvdDetail.EntryID}} = {SelectedItem!.VRD_EntryID}"
            });
        }

        public async Task PrintQCReport(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Info,
                Summary = "QC Report",
                Detail = $"Print QC Report for Lot {SelectedItem?.LotNo} — Coming Soon.",
                Duration = 3000
            });
        }

        public async Task PrintProcessInspectionReport(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null) return;

            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "RepairRejectionProcessWiseSummary.rpt",
                SelectionFormula = LastReportSql,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintReWorkRejectionChart(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null) return;

            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "RepairRejectionGraphs.rpt",
                SelectionFormula = LastReportSql,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        // ─────────────────────────────────────────────────────────────
        // Global / List Reports (Options Bar)
        // ─────────────────────────────────────────────────────────────

        private string BuildFiltersString()
        {
            var parts = new List<string>();

            if (Filter.MakerIds != null && Filter.MakerIds.Any())
                parts.Add($"Maker: {string.Join(",", Filter.MakerIds)}");
            if (Filter.ProcessIds != null && Filter.ProcessIds.Any())
                parts.Add($"Process: {string.Join(",", Filter.ProcessIds)}");
            if (Filter.ItemCatIds != null && Filter.ItemCatIds.Any())
                parts.Add($"Category: {string.Join(",", Filter.ItemCatIds)}");
            if (SelectedArticle != null)
                parts.Add($"Article: {SelectedArticle.Name}");
            if (Filter.CustomerCodes != null && Filter.CustomerCodes.Any())
                parts.Add($"Customer: {string.Join(",", Filter.CustomerCodes)}");
            if (!string.IsNullOrWhiteSpace(Filter.LotNo))
                parts.Add($"Lot No.: {Filter.LotNo}");
            if (!string.IsNullOrWhiteSpace(Filter.OrderNo))
                parts.Add($"Order No.: {Filter.OrderNo}");
            if (!string.IsNullOrWhiteSpace(Filter.MasterPONo))
                parts.Add($"Master PO: {Filter.MasterPONo}");
            if (!string.IsNullOrWhiteSpace(Filter.InActiveDays))
                parts.Add($"In-Active Days: {Filter.InActiveDays}");
            if (Filter.ShowMasterPOOnly) parts.Add("Master POs Only");
            if (Filter.ShowLastOnly) parts.Add("Last Process Done");

            return string.Join(" | ", parts);
        }

        public async Task PrintList(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "ReceivingList.rpt",
                SelectionFormula = LastReportSql,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintProcessList(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "ReceivingList_Process.rpt",
                SelectionFormula = LastReportSql,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintListSummary(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "ReceivingListSummary.rpt",
                SelectionFormula = LastReportSql,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintDeletedLots(ItemClickEventArgs args)
        {
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "DeletedLots_Receiving.rpt",
                SelectionFormula = string.Empty
            });
        }
    }
}
