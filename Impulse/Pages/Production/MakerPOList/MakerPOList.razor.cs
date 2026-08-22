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

        public MakerPOListFilter Filter { get; set; } = new MakerPOListFilter();
        public List<MakerPOListItem> AllItems { get; set; } = new List<MakerPOListItem>();
        public string ClientSearchTerm { get; set; } = string.Empty;
        public string LastReportSql { get; set; } = string.Empty;
        public bool IsLoading { get; set; } = false;
        public HashSet<long> CheckedIds { get; set; } = new HashSet<long>();
        public MakerPOListItem? SelectedItem { get; set; } = null;

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
        public bool CanCloseMakerPO => SelectedItem != null && string.IsNullOrWhiteSpace(SelectedItem.MasterPONo);
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

            int deletedCount = 0;
            int skippedCount = 0;

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
                        skippedCount++;
                        continue;
                    }
                }

                if (item.Authorized)
                {
                    int rcvCount = await MakerPOListService.CheckReceivingExistsAsync(item.EntryID);
                    if (rcvCount > 0)
                    {
                        skippedCount++;
                        continue;
                    }
                }

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
                Detail = $"Selected: {targetIds.Count}, Deleted: {deletedCount}, Skipped: {skippedCount}",
                Duration = 5000
            });

            await LoadDataAsync();
        }

        public async Task CloseMakerPO(ItemClickEventArgs args)
        {
            ResolveRowItem(args);

            if (SelectedItem == null) return;
            if (!string.IsNullOrWhiteSpace(SelectedItem.MasterPONo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Cannot Close",
                    Detail = "Only non-Master PO orders can be closed here.",
                    Duration = 4000
                });
                return;
            }

            bool success = await MakerPOListService.CloseMakerPOAsync(SelectedItem.EntryID);
            if (success)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "PO Closed",
                    Detail = $"Order #{SelectedItem.RecieptID} marked closed.",
                    Duration = 4000
                });
                await LoadDataAsync();
            }
        }

        public void EditPromises(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || string.IsNullOrWhiteSpace(SelectedItem.MasterPONo)) return;
            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Info,
                Summary = "Edit Promises",
                Detail = $"Editing promises for Master PO #{SelectedItem.MasterPONo}.",
                Duration = 3000
            });
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
            if (SelectedItem == null) return;
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

        public async Task PrintMasterPOOffice(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || string.IsNullOrWhiteSpace(SelectedItem.MasterPONo)) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssList.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem.MasterPONo}'",
                FormulaValues = new Dictionary<string, object> { { "Copy", "'OFFICE COPY'" } }
            });
        }

        public async Task PrintMasterPOMaker(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || string.IsNullOrWhiteSpace(SelectedItem.MasterPONo)) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssList.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem.MasterPONo}'",
                FormulaValues = new Dictionary<string, object> { { "Copy", "'MAKER COPY'" } }
            });
        }

        public async Task PrintMasterPOAccounts(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || string.IsNullOrWhiteSpace(SelectedItem.MasterPONo)) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssList.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem.MasterPONo}'",
                FormulaValues = new Dictionary<string, object> { { "Copy", "'ACCOUNTS COPY'" } }
            });
        }

        public async Task PrintMasterPOHideRate(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || string.IsNullOrWhiteSpace(SelectedItem.MasterPONo)) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssList.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem.MasterPONo}'",
                FormulaValues = new Dictionary<string, object> { { "HideRate", true } }
            });
        }

        public async Task PrintMasterPOStatus(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || string.IsNullOrWhiteSpace(SelectedItem.MasterPONo)) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MasterPOStatus.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem.MasterPONo}'"
            });
        }

        public async Task PrintItemPictures(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (SelectedItem == null || string.IsNullOrWhiteSpace(SelectedItem.MasterPONo)) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssListArticlePic.rpt",
                SelectionFormula = $"{{VendIssued.MasterPONo}} = '{SelectedItem.MasterPONo}'"
            });
        }

        public async Task PrintPTC(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasValidLotNo) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQEL.rpt",
                Parameters = new Dictionary<string, object>
                {
                    { "@LotNo", SelectedItem!.LotNo }
                },
                FormulaValues = new Dictionary<string, object>
                {
                    { "ComputerName", $"'IMPULSE-WEB'" }
                }
            });
        }

        public async Task PrintPTCMini(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasValidLotNo) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQEL_Mini.rpt",
                Parameters = new Dictionary<string, object>
                {
                    { "@LotNo", SelectedItem!.LotNo }
                },
                FormulaValues = new Dictionary<string, object>
                {
                    { "ComputerName", $"'IMPULSE-WEB'" }
                }
            });
        }

        public async Task PrintPTCWithPrice(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasValidLotNo) return;
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQELWithPrice.rpt",
                Parameters = new Dictionary<string, object>
                {
                    { "@LotNo", SelectedItem!.LotNo }
                },
                FormulaValues = new Dictionary<string, object>
                {
                    { "ComputerName", $"'IMPULSE-WEB'" }
                }
            });
        }

        public async Task PrintOrdersOfPurchaser(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (!HasPurchaserSelected) return;

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

            return string.Join(" | ", parts);
        }

        public async Task PrintThisList(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssuanceList.rpt",
                SelectionFormula = string.Empty,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintBatchWiseList(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "IssuanceList_BatchNoWise.rpt",
                SelectionFormula = string.Empty,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintMakerIssuanceReport(ItemClickEventArgs args)
        {
            bool hasRight = await MakerPOListService.GetUserRightAsync("MIL_Print_Maker_Issuance_Report_Valuewise", "Admin");
            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerIssuanceReportValuewise.rpt",
                SelectionFormula = string.Empty,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" }
                }
            });
        }

        public async Task PrintMakerList(ItemClickEventArgs args)
        {
            string filtersStr = BuildFiltersString();
            string dateRangeStr = $"{Filter.DtFrom:dd-MMM-yyyy} to {Filter.DtTo:dd-MMM-yyyy}";

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerList_Issuance.rpt",
                SelectionFormula = string.Empty,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"'{filtersStr}'" },
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
