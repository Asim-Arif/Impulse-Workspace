using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Http;
using DataAccessLibrary.Models.ViewModels.Production;
using DataAccessLibrary.Interface.Setup;
using Impulse.Services.Production;
using Impulse.Services;

namespace Impulse.Pages.Production.MakerIssuanceFromSF
{
    public partial class MakerIssuanceFromSF
    {
        [Inject]
        public IMakerIssuanceFromSFService SFIssuanceService { get; set; } = default!;

        [Inject]
        public IMakerPOService MakerPOService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Inject]
        public IReportNavigationService ReportNavigationService { get; set; } = default!;

        [Inject]
        public AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;

        [Inject]
        public IHttpContextAccessor HttpContextAccessor { get; set; } = default!;

        [Inject]
        public IUserDataAccess UserDataAccess { get; set; } = default!;

        [Inject]
        public IUserRoleDataAccess UserRoleDataAccess { get; set; } = default!;

        public bool IsDirectorOrAdmin { get; set; } = false;
        public string CurrentUserName { get; set; } = "System";

        [Parameter]
        public string? OrderNo { get; set; }

        [SupplyParameterFromQuery(Name = "orderNo")]
        public string? QueryOrderNo { get; set; }

        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        public string? EffectiveOrderNo => !string.IsNullOrWhiteSpace(OrderNo) ? OrderNo : QueryOrderNo;

        public bool IsPlannedOrderMode { get; set; } = false;
        public List<PlannedSFLotAllocationModel> PlannedAllocations { get; set; } = new();
        public PlannedSFLotAllocationModel? SelectedPlannedAllocation { get; set; }
        public int TotalPlannedRemaining => PlannedAllocations.Sum(p => p.RemainingToIssue);

        public void NavigateBack()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
        }

        public CreateSFIssuanceHeaderModel Header { get; set; } = new CreateSFIssuanceHeaderModel();

        public List<ProcessPOLookupModel> Processes { get; set; } = new List<ProcessPOLookupModel>();
        private ProcessPOLookupModel? _selectedProcess;
        public ProcessPOLookupModel? SelectedProcess
        {
            get => _selectedProcess;
            set
            {
                if (_selectedProcess != value)
                {
                    _selectedProcess = value;
                    _ = OnProcessChangedAsync(value);
                }
            }
        }

        private ProcessPOLookupModel? _selectedReturnProcess;
        public ProcessPOLookupModel? SelectedReturnProcess
        {
            get => _selectedReturnProcess;
            set
            {
                if (_selectedReturnProcess != value)
                {
                    _selectedReturnProcess = value;
                    _ = OnReturnProcessChangedAsync(value);
                }
            }
        }

        public List<MakerPOLookupModel> Makers { get; set; } = new List<MakerPOLookupModel>();
        private MakerPOLookupModel? _selectedMaker;
        public MakerPOLookupModel? SelectedMaker
        {
            get => _selectedMaker;
            set
            {
                if (_selectedMaker != value)
                {
                    _selectedMaker = value;
                    _ = OnMakerChangedAsync(value);
                }
            }
        }

        public List<SFStockItemLookupModel> AvailableSFItems { get; set; } = new List<SFStockItemLookupModel>();
        private SFStockItemLookupModel? _selectedItem;
        public SFStockItemLookupModel? SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (_selectedItem != value)
                {
                    _selectedItem = value;
                    _ = OnItemChangedAsync(value);
                }
            }
        }

        public List<SFStockOpeningLocationModel> OpeningLocations { get; set; } = new List<SFStockOpeningLocationModel>();
        public SFStockOpeningLocationModel? SelectedLocation { get; set; }

        public List<UnshippedOrderLookupModel> UnshippedOrders { get; set; } = new List<UnshippedOrderLookupModel>();
        public UnshippedOrderLookupModel? SelectedOrder { get; set; }

        public List<EmployeeLookupModel> Employees { get; set; } = new List<EmployeeLookupModel>();
        public EmployeeLookupModel? SelectedIssEmp { get; set; }

        public List<StagedSFIssuanceItemModel> StagedLines { get; set; } = new List<StagedSFIssuanceItemModel>();

        public List<MakerBalanceSummaryModel> MakerBalances { get; set; } = new List<MakerBalanceSummaryModel>();
        public List<MakerHistorySummaryModel> MakerHistory { get; set; } = new List<MakerHistorySummaryModel>();

        public decimal LineIssQty { get; set; }
        public decimal LineRate { get; set; }
        public DateTime LineReturnDT { get; set; } = DateTime.Today.AddDays(30);
        public string LineRemarks { get; set; } = string.Empty;

        public bool PrintSlip { get; set; } = true;
        public bool IsSaving { get; set; } = false;
        public string ActiveTab { get; set; } = "balance";

        public decimal TotalStagedQty => StagedLines.Sum(l => l.IssQty);
        public decimal TotalStagedValue => StagedLines.Sum(l => l.Value);

        public string FormatRate(decimal rate)
        {
            return IsDirectorOrAdmin ? rate.ToString("N2") : "****";
        }

        public string FormatValue(decimal value)
        {
            return IsDirectorOrAdmin ? value.ToString("N2") : "****";
        }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                    var user = authState.User;

                    bool isRoleMatch = user.IsInRole("Director") || user.IsInRole("Admin") || user.IsInRole("Administrator") || user.IsInRole("SuperAdmin");
                    bool isNameMatch = string.Equals(CurrentUserName, "admin", StringComparison.OrdinalIgnoreCase) || string.Equals(CurrentUserName, "administrator", StringComparison.OrdinalIgnoreCase);

                    if (isRoleMatch || isNameMatch)
                    {
                        IsDirectorOrAdmin = true;
                    }
                    else
                    {
                        var dbUser = await UserDataAccess.GetUserByUserNameAsync(CurrentUserName);
                        if (dbUser != null)
                        {
                            var userRoles = await UserRoleDataAccess.GetRolesByUserIdAsync(dbUser.UserID);
                            if (userRoles != null && userRoles.Any(r => r.Equals("Director", StringComparison.OrdinalIgnoreCase) || r.Equals("Admin", StringComparison.OrdinalIgnoreCase) || r.Equals("Administrator", StringComparison.OrdinalIgnoreCase)))
                            {
                                IsDirectorOrAdmin = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking user roles in MakerIssuanceFromSF: {ex.Message}");
            }

            await LoadInitialLookupsAsync();

            if (!string.IsNullOrWhiteSpace(EffectiveOrderNo))
            {
                await SetupPlannedOrderModeAsync(EffectiveOrderNo);
            }
        }

        protected override async Task OnParametersSetAsync()
        {
            if (!string.IsNullOrWhiteSpace(EffectiveOrderNo) &&
                (!IsPlannedOrderMode || !string.Equals(SelectedOrder?.OrderNo, EffectiveOrderNo, StringComparison.OrdinalIgnoreCase)))
            {
                if (Processes.Any())
                {
                    await SetupPlannedOrderModeAsync(EffectiveOrderNo);
                }
            }
        }

        public async Task SetupPlannedOrderModeAsync(string orderNo)
        {
            try
            {
                PlannedAllocations = await SFIssuanceService.GetPlannedSFAllocationsForOrderAsync(orderNo);
                if (PlannedAllocations.Any())
                {
                    IsPlannedOrderMode = true;

                    // 1. Lock/Pre-select Order
                    SelectedOrder = new UnshippedOrderLookupModel
                    {
                        OrderNo = orderNo,
                        InternalRefNo = ""
                    };
                    UnshippedOrders = new List<UnshippedOrderLookupModel> { SelectedOrder };

                    // 2. Filter available SF items to ONLY the planned items
                    var plannedItemIds = PlannedAllocations.Select(p => p.ItemID).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    var matchedItems = AvailableSFItems
                        .Where(i => plannedItemIds.Contains(i.ItemID, StringComparer.OrdinalIgnoreCase))
                        .ToList();

                    foreach (var alloc in PlannedAllocations)
                    {
                        if (!matchedItems.Any(i => i.ItemID.Equals(alloc.ItemID, StringComparison.OrdinalIgnoreCase)))
                        {
                            matchedItems.Add(new SFStockItemLookupModel
                            {
                                ItemID = alloc.ItemID,
                                ItemName = alloc.ItemName
                            });
                        }
                    }
                    AvailableSFItems = matchedItems;

                    // 3. Pre-select active allocation (prefer remaining > 0)
                    var activeAlloc = PlannedAllocations.FirstOrDefault(p => p.RemainingToIssue > 0)
                                      ?? PlannedAllocations.First();
                    SelectedPlannedAllocation = activeAlloc;

                    await ApplyPlannedAllocationAsync(activeAlloc);

                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Info,
                        Summary = "PPC Plan Mode Active",
                        Detail = $"Order #{orderNo} planned SF stock pre-selected. [{activeAlloc.ItemID}] in {activeAlloc.ProcessName} at {activeAlloc.StoreLocation} ({activeAlloc.RemainingToIssue} pcs remaining).",
                        Duration = 6000
                    });
                }
                else
                {
                    IsPlannedOrderMode = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting up planned order mode: {ex.Message}");
            }
        }

        public async Task ApplyPlannedAllocationAsync(PlannedSFLotAllocationModel alloc)
        {
            SelectedPlannedAllocation = alloc;

            // 1. Item
            _selectedItem = AvailableSFItems.FirstOrDefault(i => i.ItemID.Equals(alloc.ItemID, StringComparison.OrdinalIgnoreCase))
                            ?? new SFStockItemLookupModel { ItemID = alloc.ItemID, ItemName = alloc.ItemName };
            Header.ItemID = _selectedItem.ItemID;

            // 2. Process
            _selectedProcess = Processes.FirstOrDefault(p => p.ProcessID == alloc.ProcessID);
            _selectedReturnProcess = _selectedProcess;
            if (_selectedProcess != null)
            {
                Header.ProcessID = _selectedProcess.ProcessID;
                Makers = await MakerPOService.GetMakersForProcessAsync(_selectedProcess.ProcessID);
            }

            // 3. Stock Location - restrict strictly to the planned lot & store
            var allLocations = await SFIssuanceService.GetStockOpeningLocationsForItemAsync(alloc.ItemID, alloc.ProcessID);
            OpeningLocations = allLocations.Where(l => l.EntryID == alloc.SOO_RefID).ToList();
            if (!OpeningLocations.Any())
            {
                OpeningLocations = new List<SFStockOpeningLocationModel>
                {
                    new SFStockOpeningLocationModel
                    {
                        EntryID = alloc.SOO_RefID,
                        ItemID = alloc.ItemID,
                        ProcessID = alloc.ProcessID,
                        Shelf_RefID = alloc.Shelf_RefID,
                        StoreName = alloc.StoreLocation,
                        LotNo = alloc.LotNo,
                        InHandQty = alloc.AllocatedQty
                    }
                };
            }
            SelectedLocation = OpeningLocations.FirstOrDefault(l => l.EntryID == alloc.SOO_RefID) ?? OpeningLocations.FirstOrDefault();

            // 4. Quantity to issue
            var alreadyStaged = StagedLines.Where(l => l.OpeningEntryID == alloc.SOO_RefID && l.OrderNo == alloc.OrderNo).Sum(l => l.IssQty);
            LineIssQty = Math.Max(0, alloc.RemainingToIssue - alreadyStaged);
            LineRemarks = $"PPC Plan Order #{alloc.OrderNo} - Lot {alloc.LotNo} (Batch: {alloc.BatchNo})";

            await RefreshLineRateAsync();
            await InvokeAsync(StateHasChanged);
        }

        public async Task OnPlannedAllocationSelectChange(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int allocId))
            {
                var alloc = PlannedAllocations.FirstOrDefault(p => p.EntryID == allocId);
                if (alloc != null)
                {
                    await ApplyPlannedAllocationAsync(alloc);
                }
            }
        }

        private async Task LoadInitialLookupsAsync()
        {
            try
            {
                Processes = await SFIssuanceService.GetAllProcessesAsync();
                Employees = await MakerPOService.GetEmployeesAsync();
                AvailableSFItems = await SFIssuanceService.GetAvailableSFItemsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        public async Task HandleLotSearchKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await SearchLotNoAsync();
            }
        }

        public async Task SearchLotNoAsync()
        {
            if (string.IsNullOrWhiteSpace(Header.SearchedLotNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Input Required",
                    Detail = "Please enter a valid Lot No.",
                    Duration = 4000
                });
                return;
            }

            try
            {
                var lotResult = await SFIssuanceService.LookupLotNoAsync(Header.SearchedLotNo.Trim());
                if (lotResult == null)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Lot Not Found",
                        Detail = $"No active stock balance found for Lot No: {Header.SearchedLotNo}",
                        Duration = 5000
                    });
                    return;
                }

                // Auto-set process & item
                SelectedProcess = Processes.FirstOrDefault(p => p.ProcessID == lotResult.ProcessID);
                SelectedReturnProcess = SelectedProcess;
                SelectedItem = AvailableSFItems.FirstOrDefault(i => i.ItemID.Equals(lotResult.ItemCode, StringComparison.OrdinalIgnoreCase))
                               ?? new SFStockItemLookupModel { ItemID = lotResult.ItemCode, ItemName = lotResult.ItemName };

                OpeningLocations = await SFIssuanceService.GetStockOpeningLocationsForItemAsync(lotResult.ItemCode, lotResult.ProcessID);
                SelectedLocation = OpeningLocations.FirstOrDefault(l => l.EntryID == lotResult.SOO_RefID)
                                   ?? OpeningLocations.FirstOrDefault();

                UnshippedOrders = await SFIssuanceService.GetUnshippedOrdersForItemAsync(lotResult.ItemCode);
                if (UnshippedOrders.Any())
                {
                    SelectedOrder = UnshippedOrders.First();
                }

                LineIssQty = lotResult.SplitQty;
                LineRemarks = lotResult.Remarks;
                await RefreshLineRateAsync();

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Lot Loaded",
                    Detail = $"Lot [{lotResult.LotNo}] loaded. Process: {lotResult.ProcessName}, Item: {lotResult.ItemCode}",
                    Duration = 4000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Searching Lot",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task OnProcessChangedAsync(ProcessPOLookupModel? process)
        {
            _selectedMaker = null;
            Makers.Clear();
            if (!IsPlannedOrderMode)
            {
                SelectedItem = null;
            }

            if (process != null)
            {
                Header.ProcessID = process.ProcessID;
                if (_selectedReturnProcess == null)
                {
                    _selectedReturnProcess = process;
                }
                Makers = await MakerPOService.GetMakersForProcessAsync(process.ProcessID);
            }

            if (!IsPlannedOrderMode)
            {
                int procId = SelectedProcess?.ProcessID ?? 0;
                long vendId = SelectedMaker?.VendID ?? 0;
                int retProcId = SelectedReturnProcess?.ProcessID ?? 0;
                AvailableSFItems = await SFIssuanceService.GetAvailableSFItemsAsync(procId, vendId, retProcId);
            }

            await RefreshLineRateAsync();
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnReturnProcessChangedAsync(ProcessPOLookupModel? returnProcess)
        {
            if (!IsPlannedOrderMode)
            {
                int procId = SelectedProcess?.ProcessID ?? 0;
                long vendId = SelectedMaker?.VendID ?? 0;
                int retProcId = returnProcess?.ProcessID ?? 0;
                AvailableSFItems = await SFIssuanceService.GetAvailableSFItemsAsync(procId, vendId, retProcId);
            }

            await RefreshLineRateAsync();
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnMakerChangedAsync(MakerPOLookupModel? maker)
        {
            MakerBalances.Clear();
            MakerHistory.Clear();

            if (!IsPlannedOrderMode)
            {
                SelectedItem = null;
            }

            if (maker != null)
            {
                Header.VendID = maker.VendID;
                MakerBalances = await MakerPOService.GetMakerStockBalancesAsync(maker.VendID);
                MakerHistory = await MakerPOService.GetMakerIssuanceHistoryAsync(maker.VendID);
            }

            if (!IsPlannedOrderMode)
            {
                int procId = SelectedProcess?.ProcessID ?? 0;
                long vendId = SelectedMaker?.VendID ?? 0;
                int retProcId = SelectedReturnProcess?.ProcessID ?? 0;
                AvailableSFItems = await SFIssuanceService.GetAvailableSFItemsAsync(procId, vendId, retProcId);
            }

            await RefreshLineRateAsync();
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnItemChangedAsync(SFStockItemLookupModel? item)
        {
            OpeningLocations.Clear();
            SelectedLocation = null;

            if (item != null)
            {
                Header.ItemID = item.ItemID;
                int procId = SelectedProcess?.ProcessID ?? 0;

                if (IsPlannedOrderMode)
                {
                    var itemAlloc = PlannedAllocations.FirstOrDefault(p => p.ItemID.Equals(item.ItemID, StringComparison.OrdinalIgnoreCase) && p.RemainingToIssue > 0)
                                    ?? PlannedAllocations.FirstOrDefault(p => p.ItemID.Equals(item.ItemID, StringComparison.OrdinalIgnoreCase));
                    if (itemAlloc != null)
                    {
                        await ApplyPlannedAllocationAsync(itemAlloc);
                        return;
                    }
                }

                UnshippedOrders.Clear();
                SelectedOrder = null;
                OpeningLocations = await SFIssuanceService.GetStockOpeningLocationsForItemAsync(item.ItemID, procId);
                UnshippedOrders = await SFIssuanceService.GetUnshippedOrdersForItemAsync(item.ItemID);

                if (OpeningLocations.Any())
                {
                    SelectedLocation = OpeningLocations.First();
                    LineIssQty = SelectedLocation.InHandQty;
                    LineRemarks = SelectedLocation.Remarks;
                }

                if (UnshippedOrders.Any())
                {
                    SelectedOrder = UnshippedOrders.First();
                }
            }

            await RefreshLineRateAsync();
            await InvokeAsync(StateHasChanged);
        }

        private async Task RefreshLineRateAsync()
        {
            if (SelectedMaker == null || SelectedItem == null || SelectedProcess == null)
            {
                LineRate = 0m;
                return;
            }

            try
            {
                LineRate = await SFIssuanceService.GetItemRateAsync(
                    SelectedMaker.VendID,
                    SelectedProcess.ProcessID,
                    SelectedReturnProcess?.ProcessID ?? SelectedProcess.ProcessID,
                    SelectedItem.ItemID);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching line rate: {ex.Message}");
                LineRate = 0m;
            }
        }

        public Task<IEnumerable<ProcessPOLookupModel>> SearchProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<ProcessPOLookupModel>>(Processes);

            return Task.FromResult<IEnumerable<ProcessPOLookupModel>>(
                Processes.Where(p => p.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public Task<IEnumerable<MakerPOLookupModel>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<MakerPOLookupModel>>(Makers);

            return Task.FromResult<IEnumerable<MakerPOLookupModel>>(
                Makers.Where(m => m.VenderName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                               || (m.VendID1 != null && m.VendID1.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
        }

        public Task<IEnumerable<SFStockItemLookupModel>> SearchSFItems(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<SFStockItemLookupModel>>(AvailableSFItems);

            return Task.FromResult<IEnumerable<SFStockItemLookupModel>>(
                AvailableSFItems.Where(i => i.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                         || i.ItemID.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public Task<IEnumerable<SFStockOpeningLocationModel>> SearchOpeningLocations(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<SFStockOpeningLocationModel>>(OpeningLocations);

            return Task.FromResult<IEnumerable<SFStockOpeningLocationModel>>(
                OpeningLocations.Where(loc => loc.DisplayText.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public Task<IEnumerable<UnshippedOrderLookupModel>> SearchOrders(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<UnshippedOrderLookupModel>>(UnshippedOrders);

            return Task.FromResult<IEnumerable<UnshippedOrderLookupModel>>(
                UnshippedOrders.Where(o => o.OrderNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                        || o.InternalRefNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public Task<IEnumerable<EmployeeLookupModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<EmployeeLookupModel>>(Employees);

            return Task.FromResult<IEnumerable<EmployeeLookupModel>>(
                Employees.Where(e => e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                  || (e.EmpID != null && e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
        }

        public void AddLineToDraft()
        {
            if (SelectedProcess == null || SelectedMaker == null || SelectedItem == null || SelectedLocation == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Incomplete Selection",
                    Detail = "Please select Process, Maker, Item, and Stock Location.",
                    Duration = 4000
                });
                return;
            }

            if (LineIssQty <= 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Invalid Quantity",
                    Detail = "Issuing quantity must be greater than zero.",
                    Duration = 4000
                });
                return;
            }

            if (LineIssQty > SelectedLocation.InHandQty)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Quantity Exceeds In-Hand Stock",
                    Detail = $"Entered Qty ({LineIssQty:N0}) exceeds available in-hand stock ({SelectedLocation.InHandQty:N0}).",
                    Duration = 4000
                });
                return;
            }

            if (IsPlannedOrderMode && SelectedPlannedAllocation != null)
            {
                var alreadyStaged = StagedLines.Where(l => l.OpeningEntryID == SelectedLocation.EntryID && l.OrderNo == EffectiveOrderNo).Sum(l => l.IssQty);
                var remainingAllowed = SelectedPlannedAllocation.RemainingToIssue - (int)alreadyStaged;
                if (LineIssQty > remainingAllowed)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Quantity Exceeds Plan",
                        Detail = $"Entered Qty ({LineIssQty:N0}) exceeds remaining planned allocation ({remainingAllowed:N0} pcs) for Order #{EffectiveOrderNo}.",
                        Duration = 5000
                    });
                    return;
                }
            }

            var line = new StagedSFIssuanceItemModel
            {
                LineNo = StagedLines.Count + 1,
                OpeningEntryID = SelectedLocation.EntryID,
                OrderNo = SelectedOrder?.OrderNo ?? "Stock-Order",
                ItemCode = SelectedItem.ItemID,
                ItemName = SelectedItem.ItemName,
                IssQty = LineIssQty,
                Rate = LineRate,
                ReturnDT = LineReturnDT,
                LotNo = SelectedLocation.LotNo,
                Remarks = LineRemarks,
                ReturnProcessID = SelectedReturnProcess?.ProcessID ?? SelectedProcess.ProcessID,
                ReturnProcessName = SelectedReturnProcess?.Description ?? SelectedProcess.Description
            };

            StagedLines.Add(line);
            LineIssQty = 0;
            if (!IsPlannedOrderMode)
            {
                SelectedLocation = null;
            }
            else
            {
                var alreadyStaged = StagedLines.Where(l => l.OpeningEntryID == SelectedLocation.EntryID && l.OrderNo == EffectiveOrderNo).Sum(l => l.IssQty);
                var rem = Math.Max(0, (SelectedPlannedAllocation?.RemainingToIssue ?? 0) - (int)alreadyStaged);
                LineIssQty = rem;
            }

            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Success,
                Summary = "Line Added",
                Detail = $"SF Stock Item [{line.ItemCode}] added to draft issuance.",
                Duration = 3000
            });
        }

        public void RemoveStagedLine(StagedSFIssuanceItemModel line)
        {
            StagedLines.Remove(line);
            int idx = 1;
            foreach (var l in StagedLines)
            {
                l.LineNo = idx++;
            }
            if (IsPlannedOrderMode && SelectedPlannedAllocation != null)
            {
                var alreadyStaged = StagedLines.Where(l => l.OpeningEntryID == SelectedPlannedAllocation.SOO_RefID && l.OrderNo == EffectiveOrderNo).Sum(l => l.IssQty);
                LineIssQty = Math.Max(0, SelectedPlannedAllocation.RemainingToIssue - (int)alreadyStaged);
            }
        }

        public void ClearForm()
        {
            StagedLines.Clear();
            if (!IsPlannedOrderMode)
            {
                SelectedItem = null;
                SelectedLocation = null;
                SelectedOrder = null;
                Header = new CreateSFIssuanceHeaderModel();
            }
            else if (SelectedPlannedAllocation != null)
            {
                _ = ApplyPlannedAllocationAsync(SelectedPlannedAllocation);
            }
        }

        public async Task SaveSFIssuance()
        {
            if (SelectedProcess == null || SelectedMaker == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Error",
                    Detail = "Please select Process and Maker.",
                    Duration = 4000
                });
                return;
            }

            if (!StagedLines.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Empty Issuance",
                    Detail = "Please add at least one stock item line to the draft.",
                    Duration = 4000
                });
                return;
            }

            IsSaving = true;
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                string userName = authState?.User?.Identity?.Name ?? "System";
                int userId = 1;

                var httpContext = HttpContextAccessor.HttpContext;
                string machineName = "127.0.0.1";
                if (httpContext != null)
                {
                    string? ip = httpContext.Request.Headers["X-Forwarded-For"].FirstOrDefault();
                    if (string.IsNullOrEmpty(ip))
                    {
                        ip = httpContext.Connection.RemoteIpAddress?.ToString();
                    }
                    if (!string.IsNullOrEmpty(ip))
                    {
                        machineName = ip;
                    }
                    else
                    {
                        machineName = Environment.MachineName;
                    }
                }

                Header.IssEmpID = SelectedIssEmp?.EmpID ?? "";
                Header.ReturnProcessID = SelectedReturnProcess?.ProcessID ?? SelectedProcess.ProcessID;

                var lines = StagedLines.Select(l => new CreateSFIssuanceLineModel
                {
                    OpeningEntryID = l.OpeningEntryID,
                    OrderNo = l.OrderNo,
                    ItemCode = l.ItemCode,
                    Rate = l.Rate,
                    IssQty = l.IssQty,
                    ReturnProcessID = l.ReturnProcessID,
                    ReturnDT = l.ReturnDT,
                    Priority = 1,
                    LotNo = l.LotNo,
                    Remarks = l.Remarks
                }).ToList();

                long headerId = await SFIssuanceService.SaveSFIssuanceAsync(Header, lines, userName, userId, machineName);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "SF Issuance Saved",
                    Detail = $"Semi-Finish Maker Issuance posted successfully. Header Entry ID: {headerId}",
                    Duration = 5000
                });

                if (PrintSlip)
                {
                    await ReportNavigationService.PrintReportAsync(new ReportRequest
                    {
                        ReportName = "IssSlip.rpt",
                        SelectionFormula = $"{{VendIssued.EntryID}}={headerId}"
                    });
                }

                ClearForm();

                if (IsPlannedOrderMode && !string.IsNullOrWhiteSpace(EffectiveOrderNo))
                {
                    await SetupPlannedOrderModeAsync(EffectiveOrderNo);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }
    }
}
