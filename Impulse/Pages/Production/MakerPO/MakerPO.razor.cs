using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Http;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Impulse.Services;

namespace Impulse.Pages.Production.MakerPO
{
    public partial class MakerPO
    {
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

        public CreateMakerPOHeaderModel Header { get; set; } = new CreateMakerPOHeaderModel();

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

        public List<AssignedItemLookupModel> AssignedItems { get; set; } = new List<AssignedItemLookupModel>();
        private AssignedItemLookupModel? _selectedItem;
        public AssignedItemLookupModel? SelectedItem
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

        public List<EmployeeLookupModel> Employees { get; set; } = new List<EmployeeLookupModel>();
        public EmployeeLookupModel? SelectedIssEmp { get; set; }
        public EmployeeLookupModel? SelectedCountedBy { get; set; }
        public EmployeeLookupModel? SelectedCheckerEmp { get; set; }

        public List<SteelTypeLookupModel> SteelTypes { get; set; } = new List<SteelTypeLookupModel>();
        public SteelTypeLookupModel? SelectedSteelType { get; set; }

        public List<IssuableOrderItemModel> IssuableOrders { get; set; } = new List<IssuableOrderItemModel>();
        public IssuableOrderItemModel? SelectedIssuableOrder { get; set; }

        public List<StagedPOItemModel> StagedLines { get; set; } = new List<StagedPOItemModel>();

        public List<MakerBalanceSummaryModel> MakerBalances { get; set; } = new List<MakerBalanceSummaryModel>();
        public List<MakerHistorySummaryModel> MakerHistory { get; set; } = new List<MakerHistorySummaryModel>();

        public decimal LineIssQty { get; set; }
        public decimal LineRate { get; set; }
        public DateTime LineReturnDT { get; set; } = DateTime.Today.AddDays(30);
        public string LineBatchNo { get; set; } = "B171";

        public bool PrintSlip { get; set; } = true;
        public bool IsLoadingIssuableOrders { get; set; } = false;
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
                Processes = await MakerPOService.GetPurchaseProcessesAsync();
                Employees = await MakerPOService.GetEmployeesAsync();
                SteelTypes = await MakerPOService.GetSteelTypesAsync();
                LineBatchNo = await MakerPOService.GetNextBatchNoAsync();
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

        private async Task OnProcessChangedAsync(ProcessPOLookupModel? process)
        {
            _selectedMaker = null;
            _selectedItem = null;
            Makers.Clear();
            AssignedItems.Clear();
            IssuableOrders.Clear();
            SelectedIssuableOrder = null;

            if (process != null)
            {
                Header.ProcessID = process.ProcessID;
                Makers = await MakerPOService.GetMakersForProcessAsync(process.ProcessID);
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnMakerChangedAsync(MakerPOLookupModel? maker)
        {
            _selectedItem = null;
            AssignedItems.Clear();
            IssuableOrders.Clear();
            SelectedIssuableOrder = null;
            MakerBalances.Clear();
            MakerHistory.Clear();

            if (SelectedProcess != null && maker != null)
            {
                Header.VendID = maker.VendID;
                AssignedItems = await MakerPOService.GetAssignedItemsForMakerAndProcessAsync(maker.VendID, SelectedProcess.ProcessID);
                MakerBalances = await MakerPOService.GetMakerStockBalancesAsync(maker.VendID);
                MakerHistory = await MakerPOService.GetMakerIssuanceHistoryAsync(maker.VendID);
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnItemChangedAsync(AssignedItemLookupModel? item)
        {
            IssuableOrders.Clear();
            SelectedIssuableOrder = null;

            if (SelectedProcess != null && item != null)
            {
                Header.ItemID = item.ItemID;
                LineRate = item.Rate;
                IsLoadingIssuableOrders = true;

                try
                {
                    IssuableOrders = await MakerPOService.GetIssuableOrderItemsAsync(SelectedProcess.ProcessID, item.ItemID);
                    if (IssuableOrders.Any())
                    {
                        SelectIssuableOrder(IssuableOrders.First());
                    }
                }
                finally
                {
                    IsLoadingIssuableOrders = false;
                }
            }
            await InvokeAsync(StateHasChanged);
        }

        public void SelectIssuableOrder(IssuableOrderItemModel order)
        {
            SelectedIssuableOrder = order;
            LineIssQty = order.CalculatedIssQty;
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

        public Task<IEnumerable<AssignedItemLookupModel>> SearchAssignedItems(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<AssignedItemLookupModel>>(AssignedItems);

            return Task.FromResult<IEnumerable<AssignedItemLookupModel>>(
                AssignedItems.Where(i => i.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                      || i.ItemID.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public Task<IEnumerable<EmployeeLookupModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<EmployeeLookupModel>>(Employees);

            return Task.FromResult<IEnumerable<EmployeeLookupModel>>(
                Employees.Where(e => e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                  || (e.EmpID != null && e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
        }

        public Task<IEnumerable<SteelTypeLookupModel>> SearchSteelTypes(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<SteelTypeLookupModel>>(SteelTypes);

            return Task.FromResult<IEnumerable<SteelTypeLookupModel>>(
                SteelTypes.Where(st => st.SteelType.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public void AddLineToPO()
        {
            if (SelectedProcess == null || SelectedMaker == null || SelectedItem == null || SelectedIssuableOrder == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Incomplete Selection",
                    Detail = "Please select Process, Maker, Item, and an Order Line.",
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

            var line = new StagedPOItemModel
            {
                LineNo = StagedLines.Count + 1,
                OrderNo = SelectedIssuableOrder.OrderNo,
                ItemCode = SelectedItem.ItemID,
                ItemName = SelectedItem.ItemName,
                IssQty = LineIssQty,
                Rate = LineRate,
                ReturnDT = LineReturnDT,
                BatchNo = LineBatchNo,
                ReturnProcessID = SelectedProcess.ProcessID,
                ReturnProcessName = SelectedProcess.Description,
                AssignedUnit = SelectedItem.Unit,
                SpecialInstructions = Header.SpecialInstructions,
                IssEmpID = SelectedIssEmp?.EmpID ?? "",
                CountedBy = SelectedCountedBy?.Name ?? ""
            };

            StagedLines.Add(line);

            // Increment batch number for next item line
            if (LineBatchNo.StartsWith("B17") && int.TryParse(LineBatchNo.Substring(3), out int bNo))
            {
                LineBatchNo = $"B17{bNo + 1}";
            }
            else if (LineBatchNo.StartsWith("B") && int.TryParse(LineBatchNo.Substring(1), out int bNoAlt))
            {
                LineBatchNo = $"B{bNoAlt + 1}";
            }

            SelectedIssuableOrder = null;
            LineIssQty = 0;

            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Success,
                Summary = "Line Added",
                Detail = $"Item [{line.ItemCode}] added to draft PO.",
                Duration = 3000
            });
        }

        public void RemoveStagedLine(StagedPOItemModel line)
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
            SelectedIssuableOrder = null;
            SelectedItem = null;
            Header = new CreateMakerPOHeaderModel();
        }

        public async Task SavePO()
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
                    Summary = "Empty PO",
                    Detail = "Please add at least one item line to the draft PO.",
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
                Header.SteelType_RefID = SelectedSteelType?.SteelID ?? 0;
                Header.CheckerEmpID = SelectedCheckerEmp?.EmpID ?? "";

                var lines = StagedLines.Select(l => new CreateMakerPOLineModel
                {
                    OrderNo = l.OrderNo,
                    ItemCode = l.ItemCode,
                    Rate = l.Rate,
                    IssQty = l.IssQty,
                    ReturnProcessID = l.ReturnProcessID,
                    ReturnDT = l.ReturnDT,
                    Priority = 0,
                    BatchNo = l.BatchNo,
                    CountedBy = SelectedCountedBy?.Name ?? "",
                    IssEmpID = Header.IssEmpID,
                    Remarks = Header.SpecialInstructions
                }).ToList();

                long headerId = await MakerPOService.SaveMakerPOAsync(Header, lines, userName, userId, machineName);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "PO Saved Successfully",
                    Detail = $"Maker PO posted successfully. Header Entry ID: {headerId}",
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
