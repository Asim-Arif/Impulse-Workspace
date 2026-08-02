using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Http;
using DataAccessLibrary.Models.ViewModels.Production;
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

        public ProcessPOLookupModel? SelectedReturnProcess { get; set; }

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

        protected override async Task OnInitializedAsync()
        {
            await LoadInitialLookupsAsync();
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
            SelectedItem = null;

            if (process != null)
            {
                Header.ProcessID = process.ProcessID;
                if (SelectedReturnProcess == null)
                {
                    SelectedReturnProcess = process;
                }
                Makers = await MakerPOService.GetMakersForProcessAsync(process.ProcessID);
            }

            int procId = SelectedProcess?.ProcessID ?? 0;
            long vendId = SelectedMaker?.VendID ?? 0;
            int retProcId = SelectedReturnProcess?.ProcessID ?? 0;
            AvailableSFItems = await SFIssuanceService.GetAvailableSFItemsAsync(procId, vendId, retProcId);

            await InvokeAsync(StateHasChanged);
        }

        private async Task OnMakerChangedAsync(MakerPOLookupModel? maker)
        {
            MakerBalances.Clear();
            MakerHistory.Clear();
            SelectedItem = null;

            if (maker != null)
            {
                Header.VendID = maker.VendID;
                MakerBalances = await MakerPOService.GetMakerStockBalancesAsync(maker.VendID);
                MakerHistory = await MakerPOService.GetMakerIssuanceHistoryAsync(maker.VendID);
            }

            int procId = SelectedProcess?.ProcessID ?? 0;
            long vendId = SelectedMaker?.VendID ?? 0;
            int retProcId = SelectedReturnProcess?.ProcessID ?? 0;
            AvailableSFItems = await SFIssuanceService.GetAvailableSFItemsAsync(procId, vendId, retProcId);

            await InvokeAsync(StateHasChanged);
        }

        private async Task OnItemChangedAsync(SFStockItemLookupModel? item)
        {
            OpeningLocations.Clear();
            UnshippedOrders.Clear();
            SelectedLocation = null;
            SelectedOrder = null;

            if (item != null)
            {
                Header.ItemID = item.ItemID;
                int procId = SelectedProcess?.ProcessID ?? 0;
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
            await InvokeAsync(StateHasChanged);
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
            SelectedLocation = null;

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
        }

        public void ClearForm()
        {
            StagedLines.Clear();
            SelectedItem = null;
            SelectedLocation = null;
            SelectedOrder = null;
            Header = new CreateSFIssuanceHeaderModel();
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
