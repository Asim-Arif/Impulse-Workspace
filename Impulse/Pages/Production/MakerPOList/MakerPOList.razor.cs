using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Http;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;

namespace Impulse.Pages.Production.MakerPOList
{
    public partial class MakerPOList : ComponentBase
    {
        [Inject] public IMakerPOListService MakerPOListService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;
        [Inject] public AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;
        [Inject] public IHttpContextAccessor HttpContextAccessor { get; set; } = default!;
        [Inject] public IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;
        [Inject] public IJSRuntime JSRuntime { get; set; } = default!;
        [Inject] public SecurityService SecurityService { get; set; } = default!;
        [Inject] public Impulse.Services.WorkflowTasks.IWorkflowTaskEngine WorkflowTaskEngine { get; set; } = default!;
        [Inject] public DataAccessLibrary.Interface.Production.IProductionDeletionDataAccess ProductionDeletionDataAccess { get; set; } = default!;

        public MakerPOListFilter Filter { get; set; } = new MakerPOListFilter();
        public List<MakerPOListItem> AllItems { get; set; } = new List<MakerPOListItem>();
        public HashSet<long> PendingDeletionIssuanceEntryIds { get; set; } = new HashSet<long>();
        public HashSet<string> PendingSkipProcessLotNos { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> PendingDeletionLotNos { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public bool ShowDeleteReasonModal { get; set; } = false;
        public string DeleteRequestReason { get; set; } = string.Empty;
        public bool IsSubmittingDeleteRequest { get; set; } = false;
        public List<MakerPOListItem> DeletingItems { get; set; } = new List<MakerPOListItem>();
        private string _clientSearchTerm = string.Empty;
        public string ClientSearchTerm
        {
            get => _clientSearchTerm;
            set
            {
                if (_clientSearchTerm != value)
                {
                    _clientSearchTerm = value;
                    currentPage = 1;
                }
            }
        }
        public string LastReportSql { get; set; } = string.Empty;
        public bool IsLoading { get; set; } = false;
        public HashSet<long> CheckedIds { get; set; } = new HashSet<long>();
        public MakerPOListItem? SelectedItem { get; set; } = null;

        // Paging Parameters
        private int currentPage = 1;
        private int pageSize = 50;
        public int CurrentPage => currentPage;
        public int PageSize => pageSize;
        public int TotalCount => FilteredItems.Count;
        public int TotalPages => Math.Max(1, (int)Math.Ceiling((double)TotalCount / pageSize));

        public List<MakerPOListItem> PagedItems => FilteredItems
            .Skip((currentPage - 1) * pageSize)
            .Take(pageSize)
            .ToList();

        public void GoToPage(int page)
        {
            if (page < 1) page = 1;
            if (page > TotalPages) page = TotalPages;
            currentPage = page;
            StateHasChanged();
        }

        public void SetPageSize(int size)
        {
            pageSize = size;
            currentPage = 1;
            StateHasChanged();
        }

        // Summary Statistics (Calculated across all filtered items)
        public decimal TotalPOValue => FilteredItems.Sum(i => i.TotalValue);
        public int TotalAuthorizedCount => FilteredItems.Count(i => i.Authorized);
        public int TotalUnAuthorizedCount => FilteredItems.Count(i => !i.Authorized);

        // Multi-select Typeahead properties
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

        private LookupItemString? _selectedPurchaser;
        public LookupItemString? SelectedPurchaser
        {
            get => _selectedPurchaser;
            set
            {
                _selectedPurchaser = value;
                Filter.PurchaserEmpId = value?.Id ?? "0";
                _ = OnFilterChanged();
            }
        }

        public List<LookupItemInt> Makers { get; set; } = new List<LookupItemInt>();
        public List<LookupItemString> Categories { get; set; } = new List<LookupItemString>();
        public List<LookupItemInt> Groups { get; set; } = new List<LookupItemInt>();
        public List<LookupItemInt> Processes { get; set; } = new List<LookupItemInt>();
        public List<LookupItemString> Customers { get; set; } = new List<LookupItemString>();
        public List<LookupItemString> AllItemsList { get; set; } = new List<LookupItemString>();
        public List<LookupItemString> AllEmployeesList { get; set; } = new List<LookupItemString>();

        public bool SelectedHasMasterPO => SelectedItem != null && !string.IsNullOrWhiteSpace(SelectedItem.MasterPONo);
        public bool SelectedHasOrderNo => SelectedItem != null && !string.IsNullOrWhiteSpace(SelectedItem.OrderNo);
        public bool CanCloseMakerPO => SelectedItem != null && !string.IsNullOrWhiteSpace(SelectedItem.MasterPONo);
        public bool HasValidLotNo => SelectedItem != null && !string.IsNullOrWhiteSpace(SelectedItem.LotNo) && SelectedItem.LotNo != "0";
        public bool HasPurchaserSelected => Filter.PurchaserEmpId != "0" && !string.IsNullOrWhiteSpace(Filter.PurchaserEmpId);

        public List<MakerPOListItem> FilteredItems
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
                    (i.Hub_Name != null && i.Hub_Name.ToLower().Contains(term)) ||
                    (i.Supervisors != null && i.Supervisors.ToLower().Contains(term)) ||
                    (i.FullArticle != null && i.FullArticle.ToLower().Contains(term)) ||
                    (i.LotNo != null && i.LotNo.ToLower().Contains(term)) ||
                    (i.MasterPONo != null && i.MasterPONo.ToLower().Contains(term)) ||
                    (i.OrderNo != null && i.OrderNo.ToLower().Contains(term)) ||
                    (i.InternalRefNo != null && i.InternalRefNo.ToLower().Contains(term))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadLookupsAsync();
            await LoadDataAsync();
        }

        private async Task LoadLookupsAsync()
        {
            try
            {
                var makersTask = MakerPOListService.GetMakersAsync();
                var categoriesTask = MakerPOListService.GetItemCategoriesAsync();
                var groupsTask = MakerPOListService.GetItemGroupsAsync();
                var processesTask = MakerPOListService.GetProcessesAsync();
                var customersTask = MakerPOListService.GetCustomersAsync();
                var itemsTask = MakerPOListService.GetItemsAsync();
                var employeesTask = MakerPOListService.GetEmployeesAsync();

                await Task.WhenAll(makersTask, categoriesTask, groupsTask, processesTask, customersTask, itemsTask, employeesTask);

                Makers = await makersTask;
                Categories = await categoriesTask;
                Groups = await groupsTask;
                Processes = await processesTask;
                Customers = await customersTask;
                AllItemsList = await itemsTask;
                AllEmployeesList = await employeesTask;
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
                var result = await MakerPOListService.GetListAsync(Filter);
                AllItems = result.Items;
                LastReportSql = result.ReportSql;
                CheckedIds.Clear();
                currentPage = 1;

                try
                {
                    var pendingIdsTask = ProductionDeletionDataAccess.GetActivePendingIssuanceEntryIdsAsync();
                    var pendingSkipLotsTask = ProductionDeletionDataAccess.GetActivePendingSkipProcessLotNosAsync();
                    var pendingDelLotsTask = ProductionDeletionDataAccess.GetActivePendingDeletionLotNosAsync();

                    await Task.WhenAll(pendingIdsTask, pendingSkipLotsTask, pendingDelLotsTask);

                    PendingDeletionIssuanceEntryIds = new HashSet<long>(await pendingIdsTask);
                    PendingSkipProcessLotNos = new HashSet<string>(await pendingSkipLotsTask, StringComparer.OrdinalIgnoreCase);
                    PendingDeletionLotNos = new HashSet<string>(await pendingDelLotsTask, StringComparer.OrdinalIgnoreCase);
                }
                catch
                {
                    // Non-critical background failure
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Maker POs",
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

        // Search methods for Typeahead multi-selects

        public Task<IEnumerable<LookupItemInt>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemInt>>(Makers);

            string query = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemInt>>(
                Makers.Where(m => m.Name.ToLower().Contains(query)).ToList()
            );
        }

        public Task<IEnumerable<LookupItemString>> SearchCategories(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemString>>(Categories);

            string query = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemString>>(
                Categories.Where(c => c.Name.ToLower().Contains(query)).ToList()
            );
        }

        public Task<IEnumerable<LookupItemInt>> SearchGroups(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemInt>>(Groups);

            string query = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemInt>>(
                Groups.Where(g => g.Name.ToLower().Contains(query)).ToList()
            );
        }

        public Task<IEnumerable<LookupItemInt>> SearchProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemInt>>(Processes);

            string query = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemInt>>(
                Processes.Where(p => p.Name.ToLower().Contains(query)).ToList()
            );
        }

        public Task<IEnumerable<LookupItemString>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemString>>(Customers);

            string query = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemString>>(
                Customers.Where(c => c.Name.ToLower().Contains(query)).ToList()
            );
        }

        public Task<IEnumerable<LookupItemString>> SearchItems(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemString>>(AllItemsList);

            string query = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemString>>(
                AllItemsList.Where(i => i.Name.ToLower().Contains(query)).ToList()
            );
        }

        public Task<IEnumerable<LookupItemString>> SearchPurchasers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<LookupItemString>>(AllEmployeesList);

            string query = searchText.Trim().ToLower();
            return Task.FromResult<IEnumerable<LookupItemString>>(
                AllEmployeesList.Where(e => e.Name.ToLower().Contains(query)).ToList()
            );
        }

        public async Task OnShowMasterPOOnlyChanged()
        {
            if (Filter.ShowMasterPOOnly)
            {
                Filter.MasterPOOpen = false;
                Filter.RepairLots = false;
            }
            await LoadDataAsync();
        }

        public async Task OnMasterPOOpenChanged()
        {
            if (Filter.MasterPOOpen)
            {
                Filter.ShowMasterPOOnly = false;
                Filter.RepairLots = false;
            }
            await LoadDataAsync();
        }

        public async Task OnRepairLotsChanged()
        {
            if (Filter.RepairLots)
            {
                Filter.ShowMasterPOOnly = false;
                Filter.MasterPOOpen = false;
                Filter.RegularLotsOnly = false;
            }
            await LoadDataAsync();
        }

        public async Task OnRegularLotsOnlyChanged()
        {
            if (Filter.RegularLotsOnly)
            {
                Filter.RepairLots = false;
            }
            await LoadDataAsync();
        }

        public async Task OnRefreshClicked()
        {
            await LoadDataAsync();
        }

        public void CheckAllRows()
        {
            foreach (var item in FilteredItems)
            {
                CheckedIds.Add(item.EntryID);
            }
        }

        public void UncheckAllRows()
        {
            CheckedIds.Clear();
        }

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            bool isChecked = e.Value is bool b && b;
            if (isChecked)
            {
                CheckAllRows();
            }
            else
            {
                UncheckAllRows();
            }
        }

        public void ToggleRowCheck(long entryId, object? value)
        {
            bool isChecked = value is bool b && b;
            if (isChecked)
            {
                CheckedIds.Add(entryId);
            }
            else
            {
                CheckedIds.Remove(entryId);
            }
        }

        public void SelectRow(MakerPOListItem item)
        {
            SelectedItem = item;
        }

        public string GetRowClass(MakerPOListItem item)
        {
            if (item.BookMarkEntryID.HasValue && item.BookMarkEntryID.Value > 0)
                return "row-bookmarked";
            if (item.Closed)
                return "row-closed";
            if (!string.IsNullOrEmpty(item.ComplaintItemID))
                return "row-complaint";
            if (item.ReWorkLot.HasValue && item.ReWorkLot.Value)
                return "row-repair";

            return string.Empty;
        }

        public async Task AuthorizeSelected()
        {
            var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
            string userName = authState.User.Identity?.Name ?? "Admin";

            bool hasRight = await MakerPOListService.GetUserRightAsync("AuthorizeIssuance", userName);
            if (!hasRight)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Access Denied",
                    Detail = "You do not have permission to Authorize Issuances.",
                    Duration = 4000
                });
                return;
            }

            var unauthSelected = FilteredItems
                .Where(i => CheckedIds.Contains(i.EntryID) && !i.Authorized)
                .Select(i => i.EntryID)
                .ToList();

            if (!unauthSelected.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Info,
                    Summary = "No Un-Authorized Orders",
                    Detail = "Please check one or more un-authorized orders to authorize.",
                    Duration = 3000
                });
                return;
            }

            string machineName = Environment.MachineName;
            bool success = await MakerPOListService.AuthorizeIssuancesAsync(unauthSelected, userName, machineName);
            if (success)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Authorized Successfully",
                    Detail = $"{unauthSelected.Count} order(s) authorized.",
                    Duration = 4000
                });
                await LoadDataAsync();
            }
        }

        public async Task ShowOptionsMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is MakerPOListItem item)
            {
                SelectedItem = item;
            }
        }

        private async Task<bool> CheckIfIssuanceLockedAsync(MakerPOListItem? item)
        {
            if (item == null) return false;

            if (!string.IsNullOrWhiteSpace(item.LotNo) && PendingSkipProcessLotNos.Contains(item.LotNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Action Blocked",
                    Detail = $"Lot #{item.LotNo} is locked because a Skip Process request is awaiting Director approval.",
                    Duration = 5000
                });
                return true;
            }

            if (PendingDeletionIssuanceEntryIds.Contains(item.EntryID) || (!string.IsNullOrWhiteSpace(item.LotNo) && PendingDeletionLotNos.Contains(item.LotNo)))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Action Blocked",
                    Detail = $"Order/Lot #{item.RecieptID} is locked because a deletion request is awaiting Director approval.",
                    Duration = 5000
                });
                return true;
            }

            bool isLocked = await ProductionDeletionDataAccess.IsIssuanceLockedAsync(item.EntryID, item.LotNo);
            if (isLocked)
            {
                PendingDeletionIssuanceEntryIds.Add(item.EntryID);
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Action Blocked",
                    Detail = $"Order/Lot #{item.RecieptID} is locked because a deletion request is awaiting Director approval.",
                    Duration = 5000
                });
                return true;
            }

            return false;
        }

        public async Task DeleteSelected(ItemClickEventArgs args)
        {
            ResolveRowItem(args);

            var targetIds = CheckedIds.Any()
                ? FilteredItems.Where(i => CheckedIds.Contains(i.EntryID)).ToList()
                : (SelectedItem != null ? new List<MakerPOListItem> { SelectedItem } : new List<MakerPOListItem>());

            if (!targetIds.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Order Selected",
                    Detail = "Please select at least one order to delete.",
                    Duration = 3000
                });
                return;
            }

            // Check if any selected item is already pending deletion
            var alreadyPending = targetIds.Where(i => PendingDeletionIssuanceEntryIds.Contains(i.EntryID)).ToList();
            if (alreadyPending.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Request Pending",
                    Detail = $"Order/Lot #{string.Join(", ", alreadyPending.Select(p => p.RecieptID))} is already awaiting Director approval.",
                    Duration = 5000
                });
                return;
            }

            // Pre-validate loans & receiving
            var validTargets = new List<MakerPOListItem>();
            foreach (var item in targetIds)
            {
                if (!string.IsNullOrWhiteSpace(item.MasterPONo))
                {
                    var (shortLoan, longLoan) = await MakerPOListService.CheckLoanExistsAsync(item.MasterPONo);
                    if (shortLoan || longLoan)
                    {
                        NotificationService.Notify(new Radzen.NotificationMessage
                        {
                            Severity = Radzen.NotificationSeverity.Error,
                            Summary = "Delete Blocked",
                            Detail = $"Cannot delete Order #{item.RecieptID}: Short/Long term loan is issued.",
                            Duration = 5000
                        });
                        continue;
                    }
                }

                if (item.Authorized)
                {
                    int rcvCount = await MakerPOListService.CheckReceivingExistsAsync(item.EntryID);
                    if (rcvCount > 0)
                    {
                        NotificationService.Notify(new Radzen.NotificationMessage
                        {
                            Severity = Radzen.NotificationSeverity.Error,
                            Summary = "Delete Blocked",
                            Detail = $"Cannot delete Order #{item.RecieptID}: Goods have already been received against this order.",
                            Duration = 5000
                        });
                        continue;
                    }
                }

                validTargets.Add(item);
            }

            if (!validTargets.Any()) return;

            var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
            string userName = authState.User.Identity?.Name ?? "Admin";
            var user = authState.User;

            bool requiresApproval = await WorkflowTaskEngine.IsApprovalRequiredAsync("DeleteProductionIssuance", userName, user);

            if (requiresApproval)
            {
                DeletingItems = validTargets;
                DeleteRequestReason = string.Empty;
                ShowDeleteReasonModal = true;
                StateHasChanged();
                return;
            }

            // Verify Password via Global Security Service for exempt users
            bool isAuthorized = await SecurityService.VerifyActionAsync("DeleteProdIss");

            if (!isAuthorized)
            {
                NotificationService.Notify(Radzen.NotificationSeverity.Error, "Unauthorized", "Incorrect password or action cancelled.");
                return;
            }

            int deletedCount = 0;
            foreach (var item in validTargets)
            {
                bool result = await MakerPOListService.DeleteIssuanceAsync(item.EntryID);
                if (result)
                {
                    deletedCount++;
                }
            }

            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = deletedCount > 0 ? Radzen.NotificationSeverity.Success : Radzen.NotificationSeverity.Warning,
                Summary = "Delete Operation Complete",
                Detail = $"Selected: {validTargets.Count}, Deleted: {deletedCount}",
                Duration = 5000
            });

            await LoadDataAsync();
        }

        public void CloseDeleteReasonModal()
        {
            ShowDeleteReasonModal = false;
            DeletingItems.Clear();
            DeleteRequestReason = string.Empty;
        }

        public async Task SubmitDeleteRequestAsync()
        {
            if (!DeletingItems.Any()) return;

            if (string.IsNullOrWhiteSpace(DeleteRequestReason))
            {
                NotificationService.Notify(Radzen.NotificationSeverity.Warning, "Validation", "Please provide a reason for the deletion request.");
                return;
            }

            IsSubmittingDeleteRequest = true;
            StateHasChanged();

            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                string userName = authState.User.Identity?.Name ?? "System";

                int submittedCount = 0;
                foreach (var item in DeletingItems)
                {
                    var request = new ProductionDeletionRequestModel
                    {
                        RequestType = "LotIssuance",
                        EntityRefID = item.EntryID,
                        LotNo = item.LotNo ?? "",
                        OrderNo = item.MasterPONo ?? item.OrderNo,
                        ItemCode = item.ItemID,
                        ItemName = item.ItemName,
                        ProcessID = item.ProcessID,
                        ProcessName = item.Description,
                        MakerID = item.VendID,
                        MakerName = item.VenderName,
                        Qty = item.TotalIssQty,
                        RequestedBy = userName,
                        RequestedDT = DateTime.Now,
                        Reason = DeleteRequestReason.Trim(),
                        MachineName = Environment.MachineName,
                        Status = "Pending"
                    };

                    await WorkflowTaskEngine.RequestProductionIssuanceDeletionAsync(request);
                    submittedCount++;
                }

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Requests Submitted",
                    Detail = $"{submittedCount} issuance deletion request(s) submitted to Director for approval.",
                    Duration = 5000
                });

                ShowDeleteReasonModal = false;
                DeletingItems.Clear();
                DeleteRequestReason = string.Empty;
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Submission Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSubmittingDeleteRequest = false;
                StateHasChanged();
            }
        }

        public async Task CloseMakerPO(ItemClickEventArgs args)
        {
            ResolveRowItem(args);

            if (SelectedItem == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Selection",
                    Detail = "Please select an order first.",
                    Duration = 3000
                });
                return;
            }

            if (await CheckIfIssuanceLockedAsync(SelectedItem)) return;

            if (!string.IsNullOrWhiteSpace(SelectedItem.LotNo) && SelectedItem.LotNo != "0")
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Not a PO",
                    Detail = "This is a Lot, not a Purchase Order.",
                    Duration = 4000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(SelectedItem.MasterPONo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Cannot Close",
                    Detail = "Only Master POs can be closed.",
                    Duration = 4000
                });
                return;
            }

            bool isAuthorized = await SecurityService.VerifyActionAsync("CloseMakerPO");
            if (!isAuthorized)
            {
                return;
            }

            bool success = await MakerPOListService.CloseMakerPOAsync(SelectedItem.EntryID, SelectedItem.MasterPONo);
            if (success)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Master PO Closed",
                    Detail = $"Master PO #{SelectedItem.MasterPONo} marked closed.",
                    Duration = 4000
                });
                await LoadDataAsync();
            }
        }

        public bool IsEditPromisesModalOpen { get; set; } = false;
        public bool IsLoadingPromises { get; set; } = false;
        public bool IsSavingPromises { get; set; } = false;
        public List<MakerPOReturnDateItemDto> ReturnDatesList { get; set; } = new();
        public MakerPOListItem? SelectedItemForPromises { get; set; }

        public async Task EditPromises(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Selection",
                    Detail = "Please select an order first.",
                    Duration = 3000
                });
                return;
            }

            SelectedItemForPromises = SelectedItem;
            IsEditPromisesModalOpen = true;
            IsLoadingPromises = true;

            try
            {
                ReturnDatesList = await MakerPOListService.GetMakerPOReturnDatesAsync(SelectedItem.EntryID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Return Dates",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoadingPromises = false;
            }
        }

        public void CloseEditPromisesModal()
        {
            IsEditPromisesModalOpen = false;
            ReturnDatesList.Clear();
            SelectedItemForPromises = null;
        }

        public async Task SavePromisesAsync()
        {
            if (SelectedItemForPromises == null || !ReturnDatesList.Any())
            {
                CloseEditPromisesModal();
                return;
            }

            IsSavingPromises = true;
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                string userName = authState.User.Identity?.Name ?? "Admin";
                string machineName = Environment.MachineName;

                bool success = await MakerPOListService.UpdateMakerPOReturnDatesAsync(
                    SelectedItemForPromises.EntryID,
                    ReturnDatesList,
                    userName,
                    machineName
                );

                if (success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Return Dates Updated",
                        Detail = $"Return dates updated successfully for Master PO #{SelectedItemForPromises.MasterPONo}.",
                        Duration = 4000
                    });

                    CloseEditPromisesModal();
                    await LoadDataAsync();
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Update Incomplete",
                        Detail = "No records were updated.",
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Updating Return Dates",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSavingPromises = false;
            }
        }

        public void AddBookmark(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null) return;
            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Info,
                Summary = "Bookmark",
                Detail = $"Bookmarked Order #{SelectedItem.RecieptID}.",
                Duration = 3000
            });
        }

        public void AddFollowup(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Selection",
                    Detail = "Please select an order first.",
                    Duration = 3000
                });
                return;
            }
            if (string.IsNullOrWhiteSpace(SelectedItem.OrderNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Order #",
                    Detail = "The selected order does not have an Order #.",
                    Duration = 4000
                });
                return;
            }

            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Info,
                Summary = "Followup",
                Detail = $"Followup entry for Order #{SelectedItem.OrderNo}.",
                Duration = 3000
            });
        }

        // ────── Row-Specific Reports ──────

        private List<long> GetTargetEntryIds()
        {
            if (CheckedIds.Any())
                return CheckedIds.ToList();
            if (SelectedItem != null)
                return new List<long> { SelectedItem.EntryID };

            return new List<long>();
        }

        public async Task PrintIssSlip(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            var ids = GetTargetEntryIds();
            if (!ids.Any()) return;

            string idList = string.Join(",", ids);
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssSlip.rpt",
                SelectionFormula = $"{{VendIssued.EntryID}} IN [{idList}]"
            });
        }

        public async Task PrintIssSlipNoRate(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            var ids = GetTargetEntryIds();
            if (!ids.Any()) return;

            string idList = string.Join(",", ids);
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssSlip.rpt",
                SelectionFormula = $"{{VendIssued.EntryID}} IN [{idList}]",
                FormulaValues = new Dictionary<string, object> { { "HideCustomer", true } }
            });
        }

        public async Task PrintIssSlipMini(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            var ids = GetTargetEntryIds();
            if (!ids.Any()) return;

            string idList = string.Join(",", ids);
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssSlipMin.rpt",
                SelectionFormula = $"{{VendIssued.EntryID}} IN [{idList}]"
            });
        }

        public async Task PrintTag(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            var ids = GetTargetEntryIds();
            if (!ids.Any()) return;

            string idList = string.Join(",", ids);
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssSlipTag.rpt",
                SelectionFormula = $"{{VendIssued.EntryID}} IN [{idList}]"
            });
        }

        public async Task PrintInternal(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            var ids = GetTargetEntryIds();
            if (!ids.Any()) return;

            string idList = string.Join(",", ids);
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssSlipInternal.rpt",
                SelectionFormula = $"{{VendIssued.EntryID}} IN [{idList}]"
            });
        }

        private bool CheckHasMasterPO()
        {
            if (SelectedItem == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Selection",
                    Detail = "Please select an order first.",
                    Duration = 3000
                });
                return false;
            }
            if (string.IsNullOrWhiteSpace(SelectedItem.MasterPONo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Master PO",
                    Detail = "The selected order does not have a Master PO #.",
                    Duration = 4000
                });
                return false;
            }
            return true;
        }

        public async Task PrintMasterPOOffice(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!CheckHasMasterPO()) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssList.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem!.MasterPONo}'",
                FormulaValues = new Dictionary<string, object> { { "Copy", "'OFFICE COPY'" } }
            });
        }

        public async Task PrintMasterPOMaker(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!CheckHasMasterPO()) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssList.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem!.MasterPONo}'",
                FormulaValues = new Dictionary<string, object> { { "Copy", "'MAKER COPY'" } }
            });
        }

        public async Task PrintMasterPOAccounts(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!CheckHasMasterPO()) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssList.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem!.MasterPONo}'",
                FormulaValues = new Dictionary<string, object> { { "Copy", "'ACCOUNTS COPY'" } }
            });
        }

        public async Task PrintMasterPOHideRate(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!CheckHasMasterPO()) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssList.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem!.MasterPONo}'",
                FormulaValues = new Dictionary<string, object> { { "HideRate", true } }
            });
        }

        public async Task PrintMasterPOStatus(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!CheckHasMasterPO()) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MasterPOStatus.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem!.MasterPONo}'"
            });
        }

        public async Task PrintItemPictures(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!CheckHasMasterPO()) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssListArticlePic.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem!.MasterPONo}'"
            });
        }

        public async Task PrintPTC(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Selection",
                    Detail = "Please select an order first.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(SelectedItem.LotNo) || SelectedItem.LotNo == "0")
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Not a Lot",
                    Detail = "This is a PO, not a Lot.",
                    Duration = 4000
                });
                return;
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQel.rpt",
                Parameters = new Dictionary<string, object>
                {
                    { "@LotNo", SelectedItem.LotNo }
                }
            });
        }

        public async Task PrintPTCMini(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Selection",
                    Detail = "Please select an order first.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(SelectedItem.LotNo) || SelectedItem.LotNo == "0")
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Not a Lot",
                    Detail = "This is a PO, not a Lot.",
                    Duration = 4000
                });
                return;
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQEL_Mini.rpt",
                Parameters = new Dictionary<string, object>
                {
                    { "@LotNo", SelectedItem.LotNo }
                }
            });
        }

        public async Task PrintPTCWithPrice(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Selection",
                    Detail = "Please select an order first.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(SelectedItem.LotNo) || SelectedItem.LotNo == "0")
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Not a Lot",
                    Detail = "This is a PO, not a Lot.",
                    Duration = 4000
                });
                return;
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQELWithPrice.rpt",
                Parameters = new Dictionary<string, object>
                {
                    { "@LotNo", SelectedItem.LotNo }
                }
            });
        }

        public async Task PrintOrdersOfPurchaser(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasPurchaserSelected)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Purchaser Required",
                    Detail = "Please select a Purchaser in the filter first.",
                    Duration = 4000
                });
                return;
            }

            string formula = $"{{VVendIssdDetail_Simple.VID_EmpID}} = '{Filter.PurchaserEmpId}' AND {{VVendIssdDetail_Simple.DT}} in Date({Filter.DtFrom.Year}, {Filter.DtFrom.Month}, {Filter.DtFrom.Day}) to Date({Filter.DtTo.Year}, {Filter.DtTo.Month}, {Filter.DtTo.Day})";
            if (!string.IsNullOrWhiteSpace(Filter.OrderNo))
            {
                formula += $" AND {{FCustomerOrders.OrderNo}} = '{Filter.OrderNo.Trim()}'";
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PurchaserwiseMonthlyReport.rpt",
                SelectionFormula = formula
            });
        }

        public async Task PrintPurchasePlan(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            string formula = $"{{VVendIssdDetail_Simple.DT}} in Date({Filter.DtFrom.Year}, {Filter.DtFrom.Month}, {Filter.DtFrom.Day}) to Date({Filter.DtTo.Year}, {Filter.DtTo.Month}, {Filter.DtTo.Day})";
            if (Filter.MakerIds != null && Filter.MakerIds.Any())
            {
                formula += $" AND {{VVendIssdDetail_Simple.VendID}} IN [{string.Join(",", Filter.MakerIds)}]";
            }
            if (HasPurchaserSelected)
            {
                formula += $" AND {{VVendIssdDetail_Simple.VID_EmpID}} = '{Filter.PurchaserEmpId}'";
            }
            if (!string.IsNullOrWhiteSpace(Filter.OrderNo))
            {
                formula += $" AND {{VrptOrders_ForProduction.OrderNo}} = '{Filter.OrderNo.Trim()}'";
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PlanforPurchase.rpt",
                SelectionFormula = formula
            });
        }

        public async Task PrintPurchaseCalendar(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            string formula = $"{{DateLookup.DateFull}} in Date({Filter.DtFrom.Year}, {Filter.DtFrom.Month}, {Filter.DtFrom.Day}) to Date({Filter.DtTo.Year}, {Filter.DtTo.Month}, {Filter.DtTo.Day})";
            string makerId = Filter.MakerIds != null && Filter.MakerIds.Any() ? string.Join(",", Filter.MakerIds) : "0";
            string groupIds = Filter.ItemGroupIds != null && Filter.ItemGroupIds.Any() ? string.Join(",", Filter.ItemGroupIds) : "0";
            string custCodes = Filter.CustomerCodes != null && Filter.CustomerCodes.Any() ? string.Join(",", Filter.CustomerCodes) : "0";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PurchaseCalendar.rpt",
                SelectionFormula = formula,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Maker", $"'{makerId}'" },
                    { "ItemGroup", $"'{groupIds}'" },
                    { "Customer", $"'{custCodes}'" }
                }
            });
        }

        public async Task PrintFollowupReport(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "Followup_Report.rpt",
                SelectionFormula = SelectedItem != null ? $"{{VendIssued.EntryID}} = {SelectedItem.EntryID}" : string.Empty
            });
        }

        // ────── Options Bar Menu (Global List Reports) ──────

        private string BuildFiltersString()
        {
            List<string> parts = new List<string>();

            if (Filter.MakerIds != null && Filter.MakerIds.Any())
            {
                var makerNames = Makers.Where(m => Filter.MakerIds.Contains(m.Id)).Select(m => m.Name);
                parts.Add($"Maker: {(makerNames.Any() ? string.Join(",", makerNames) : string.Join(",", Filter.MakerIds))}");
            }
            if (Filter.ProcessIds != null && Filter.ProcessIds.Any())
            {
                var procNames = Processes.Where(p => Filter.ProcessIds.Contains(p.Id)).Select(p => p.Name);
                parts.Add($"Process: {(procNames.Any() ? string.Join(",", procNames) : string.Join(",", Filter.ProcessIds))}");
            }
            if (Filter.ItemCatIds != null && Filter.ItemCatIds.Any())
            {
                var catNames = Categories.Where(c => Filter.ItemCatIds.Contains(c.Id)).Select(c => c.Name);
                parts.Add($"Category: {(catNames.Any() ? string.Join(",", catNames) : string.Join(",", Filter.ItemCatIds))}");
            }
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
            if (Filter.ShowMasterPOOnly)
                parts.Add("Master POs Only");
            if (Filter.MasterPOOpen)
                parts.Add("Open Master POs Only");
            if (Filter.RepairLots)
                parts.Add("Repair Lots Only");

            return string.Join(" | ", parts);
        }

        private string BuildSelectionFormula(out string dateRangeStr)
        {
            var conditions = new List<string>();
            DateTime dtFrom = Filter.DtFrom;
            DateTime dtTo = Filter.DtTo;

            // 1. Due Days or Date Range
            if (Filter.DueDaysMode > 0)
            {
                int dueDays = Filter.DueDaysMode switch
                {
                    1 => 3,
                    2 => 7,
                    3 => Filter.DueDaysCustom,
                    _ => 0
                };
                DateTime targetDt = DateTime.Today.AddDays(dueDays);
                dtFrom = targetDt < DateTime.Today ? targetDt : DateTime.Today;
                dtTo = targetDt < DateTime.Today ? DateTime.Today : targetDt;

                conditions.Add($"{{VVendIssued.DT}} in Date({dueFromDate(dtFrom)}) to Date({dueToDate(dtTo)})");
            }
            else
            {
                DateTime now = DateTime.Today;
                switch (Filter.DateRangeIndex)
                {
                    case 0: dtFrom = now; dtTo = now; break;
                    case 1: dtFrom = now.AddDays(-15); dtTo = now; break;
                    case 2: dtFrom = now.AddDays(-30); dtTo = now; break;
                    case 3: dtFrom = now.AddDays(-60); dtTo = now; break;
                    case 4: dtFrom = now.AddDays(-90); dtTo = now; break;
                    case 5:
                        dtFrom = Filter.DtFrom;
                        dtTo = Filter.DtTo;
                        break;
                }

                // In legacy logic, if LotNo, OrderNo, or MasterPONo is specified, date filter is bypassed
                if (string.IsNullOrWhiteSpace(Filter.LotNo) && string.IsNullOrWhiteSpace(Filter.OrderNo) && string.IsNullOrWhiteSpace(Filter.MasterPONo))
                {
                    conditions.Add($"{{VVendIssued.DT}} in Date({dtFrom.Year}, {dtFrom.Month}, {dtFrom.Day}) to Date({dtTo.Year}, {dtTo.Month}, {dtTo.Day})");
                }
            }

            dateRangeStr = $"{dtFrom:dd-MMM-yyyy} to {dtTo:dd-MMM-yyyy}";

            // 2. Maker Filter (VendID)
            if (Filter.MakerIds != null && Filter.MakerIds.Any())
            {
                conditions.Add($"{{VVendIssued.VendID}} in [{string.Join(",", Filter.MakerIds)}]");
            }

            // 3. Item / Article Filter (ItemID)
            if (!string.IsNullOrWhiteSpace(Filter.ItemId) && Filter.ItemId != "0")
            {
                conditions.Add($"{{VVendIssued.ItemID}} = '{Filter.ItemId.Replace("'", "''")}'");
            }

            // 4. Category Filter (CatID)
            if (Filter.ItemCatIds != null && Filter.ItemCatIds.Any())
            {
                if (Filter.ItemCatIds.All(c => int.TryParse(c, out _)))
                {
                    conditions.Add($"{{VVendIssued.CatID}} in [{string.Join(",", Filter.ItemCatIds)}]");
                }
                else
                {
                    var cats = Filter.ItemCatIds.Select(c => $"'{c.Replace("'", "''")}'");
                    conditions.Add($"{{VVendIssued.CatID}} in [{string.Join(",", cats)}]");
                }
            }

            // 5. Item Group Filter (GroupID)
            if (Filter.ItemGroupIds != null && Filter.ItemGroupIds.Any())
            {
                conditions.Add($"{{VVendIssued.GroupID}} in [{string.Join(",", Filter.ItemGroupIds)}]");
            }

            // 6. Process Filter (ProcessID)
            if (Filter.ProcessIds != null && Filter.ProcessIds.Any())
            {
                conditions.Add($"{{VVendIssued.ProcessID}} in [{string.Join(",", Filter.ProcessIds)}]");
            }

            // 7. Lot No Filter (LotNo)
            if (!string.IsNullOrWhiteSpace(Filter.LotNo))
            {
                conditions.Add($"{{VVendIssued.LotNo}} = '{Filter.LotNo.Trim().Replace("'", "''")}'");
            }

            // 8. Order No Filter (OrderNo)
            if (!string.IsNullOrWhiteSpace(Filter.OrderNo))
            {
                conditions.Add($"{{VVendIssued.OrderNo}} = '{Filter.OrderNo.Trim().Replace("'", "''")}'");
            }

            // 9. Master PO No (MasterPONo)
            if (!string.IsNullOrWhiteSpace(Filter.MasterPONo))
            {
                conditions.Add($"{{VVendIssued.MasterPONo}} = '{Filter.MasterPONo.Trim().Replace("'", "''")}'");
            }

            // 10. Show Master PO Only
            if (Filter.ShowMasterPOOnly)
            {
                conditions.Add("not IsNull({VVendIssued.MasterPONo}) and {VVendIssued.MasterPONo} <> ''");
            }

            // 11. Repair Lots
            if (Filter.RepairLots)
            {
                conditions.Add("{VVendIssued.ReWorkLot} = 1");
            }

            // 12. Regular Lots Only
            if (Filter.RegularLotsOnly)
            {
                conditions.Add("{VVendIssued.ReWorkLot} = 0");
            }

            // 13. Bookmarks
            if (Filter.Bookmarks)
            {
                conditions.Add("not IsNull({VVendIssued.BookMarkEntryID})");
            }

            // 14. Customer Codes (CustCode)
            if (Filter.CustomerCodes != null && Filter.CustomerCodes.Any())
            {
                var custs = Filter.CustomerCodes.Select(c => $"'{c.Replace("'", "''")}'");
                conditions.Add($"{{VVendIssued.CustCode}} in [{string.Join(",", custs)}]");
            }

            // 15. Purchaser
            if (!string.IsNullOrWhiteSpace(Filter.PurchaserEmpId) && Filter.PurchaserEmpId != "0")
            {
                conditions.Add($"{{VVendIssued.IssEmpID}} = '{Filter.PurchaserEmpId.Replace("'", "''")}'");
            }

            return conditions.Any() ? string.Join(" and ", conditions) : string.Empty;
        }

        private static string dueFromDate(DateTime d) => $"{d.Year}, {d.Month}, {d.Day}";
        private static string dueToDate(DateTime d) => $"{d.Year}, {d.Month}, {d.Day}";

        public async Task PrintThisList(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string selectionFormula = BuildSelectionFormula(out string dateRangeStr);

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssuanceList.rpt",
                SelectionFormula = selectionFormula,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr.Replace("'", "''")}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintBatchWiseList(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string selectionFormula = BuildSelectionFormula(out string dateRangeStr);

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssuanceList_BatchNoWise.rpt",
                SelectionFormula = selectionFormula,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr.Replace("'", "''")}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintMakerIssuanceReport(ItemClickEventArgs args)
        {
            bool hasRight = await MakerPOListService.GetUserRightAsync("MIL_Print_Maker_Issuance_Report_Valuewise", "Admin");
            string filtersStr = BuildFiltersString();
            string selectionFormula = BuildSelectionFormula(out string dateRangeStr);

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerIssuanceReportValuewise.rpt",
                SelectionFormula = selectionFormula,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr.Replace("'", "''")}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintMakerList(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string selectionFormula = BuildSelectionFormula(out string dateRangeStr);

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerList_Issuance.rpt",
                SelectionFormula = selectionFormula,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr.Replace("'", "''")}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintMakerBalanceReport(ItemClickEventArgs args)
        {
            string formula = "{VendIssdDetail.RcvdQty} < {VendIssdDetail.IssQty} AND {VendIssued.Closed} = FALSE";
            if (Filter.ProcessIds != null && Filter.ProcessIds.Any())
            {
                formula += $" AND {{VendIssued.ProcessID}} IN [{string.Join(",", Filter.ProcessIds)}]";
            }
            if (Filter.MakerIds != null && Filter.MakerIds.Any())
            {
                formula += $" AND {{VendIssued.VendID}} IN [{string.Join(",", Filter.MakerIds)}]";
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerBalanceReport.rpt",
                SelectionFormula = formula
            });
        }

        public async Task PrintMakerBalanceReportSummary(ItemClickEventArgs args)
        {
            string formula = "{VendIssdDetail.RcvdQty} < {VendIssdDetail.IssQty} AND {VendIssued.Closed} = FALSE";
            if (Filter.ProcessIds != null && Filter.ProcessIds.Any())
            {
                formula += $" AND {{VendIssued.ProcessID}} IN [{string.Join(",", Filter.ProcessIds)}]";
            }
            if (Filter.MakerIds != null && Filter.MakerIds.Any())
            {
                formula += $" AND {{VendIssued.VendID}} IN [{string.Join(",", Filter.MakerIds)}]";
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerBalanceReport_Summary.rpt",
                SelectionFormula = formula
            });
        }

        public async Task PrintSFStock(ItemClickEventArgs args)
        {
            string formula = "{@NetQty} > 0";
            if (SelectedArticle != null)
            {
                formula += $" AND {{StockOrderOpening.ItemID}} = '{SelectedArticle.Id}'";
            }
            if (Filter.ProcessIds != null && Filter.ProcessIds.Any())
            {
                formula += $" AND {{StockOrderOpening.ProcessID}} IN [{string.Join(",", Filter.ProcessIds)}]";
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "SFStockReportStorewise.rpt",
                SelectionFormula = formula
            });
        }
    }
}
