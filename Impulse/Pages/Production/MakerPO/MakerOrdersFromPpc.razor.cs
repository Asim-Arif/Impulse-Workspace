using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Http;
using Radzen;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Impulse.Services;
using Impulse.Models;

namespace Impulse.Pages.Production.MakerPO
{
    public partial class MakerOrdersFromPpc : ComponentBase
    {
        [Inject] public IPpcMakerOrderService PpcMakerOrderService { get; set; } = default!;
        [Inject] public IMakerPOService MakerPOService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;
        [Inject] public IHttpContextAccessor HttpContextAccessor { get; set; } = default!;

        [Parameter] public string? OrderNo { get; set; }
        [SupplyParameterFromQuery(Name = "orderNo")] public string? QueryOrderNo { get; set; }
        [SupplyParameterFromQuery(Name = "returnUrl")] public string? ReturnUrl { get; set; }

        public string? EffectiveOrderNo => !string.IsNullOrWhiteSpace(OrderNo) ? OrderNo : QueryOrderNo;

        public bool IsLoading { get; set; } = true;
        public bool IsSaving { get; set; } = false;

        public PpcMakerPoOrderHeaderDto? OrderHeader { get; set; }
        public List<PpcMakerPoItemRowDto> PlannedLines { get; set; } = new();

        public List<PpcMakerPoOrderHeaderDto> ActiveOrdersWithPurchases { get; set; } = new();
        public PpcMakerPoOrderHeaderDto? SelectedOrderLookup { get; set; }

        public List<MakerPOLookupModel> AllMakers { get; set; } = new();
        public List<ProcessPOLookupModel> Processes { get; set; } = new();
        public List<EmployeeLookupModel> Employees { get; set; } = new();
        public List<SteelTypeLookupModel> SteelTypes { get; set; } = new();

        public EmployeeLookupModel? SelectedIssEmp { get; set; }
        public EmployeeLookupModel? SelectedCountedBy { get; set; }
        
        private int? _selectedSteelTypeId;
        public int? SelectedSteelTypeId
        {
            get => _selectedSteelTypeId;
            set
            {
                _selectedSteelTypeId = value;
                if (value.HasValue && PlannedLines != null)
                {
                    foreach (var line in PlannedLines.Where(l => !l.IsPosted && !l.SteelType_RefID.HasValue))
                    {
                        line.SteelType_RefID = value.Value;
                    }
                }
            }
        }
        public SteelTypeLookupModel? SelectedSteelType => SteelTypes.FirstOrDefault(s => s.SteelID == SelectedSteelTypeId);

        public string CurrentUserName { get; set; } = "Purchaser";
        public int CurrentUserId { get; set; } = 1;
        public string CurrentMachineName { get; set; } = "WEB-CLIENT";

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoading = true;

                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                var user = authState.User;
                if (user.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = user.Identity.Name ?? "Purchaser";
                    var idClaim = user.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                    if (int.TryParse(idClaim, out int parsedId)) CurrentUserId = parsedId;
                }

                CurrentMachineName = HttpContextAccessor.HttpContext?.Connection?.RemoteIpAddress?.ToString() ?? "WEB-CLIENT";

                // Load reference datasets
                AllMakers = await PpcMakerOrderService.GetAllMakersAsync();
                Processes = await MakerPOService.GetPurchaseProcessesAsync();
                Employees = await MakerPOService.GetEmployeesAsync();
                SteelTypes = await MakerPOService.GetSteelTypesAsync();
                ActiveOrdersWithPurchases = await PpcMakerOrderService.GetActiveOrdersWithPendingPurchasesAsync();

                if (Employees.Any())
                {
                    SelectedIssEmp = Employees.FirstOrDefault();
                    SelectedCountedBy = Employees.FirstOrDefault();
                }

                if (SteelTypes.Any())
                {
                    SelectedSteelTypeId = SteelTypes.FirstOrDefault()?.SteelID;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Initialization Error",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected override async Task OnParametersSetAsync()
        {
            if (!string.IsNullOrWhiteSpace(EffectiveOrderNo))
            {
                await LoadOrderDataAsync(EffectiveOrderNo);
            }
            else
            {
                OrderHeader = null;
                PlannedLines.Clear();
                SelectedOrderLookup = null;
            }
        }

        public async Task LoadOrderDataAsync(string orderNo)
        {
            try
            {
                IsLoading = true;
                OrderHeader = await PpcMakerOrderService.GetOrderHeaderAsync(orderNo);
                PlannedLines = await PpcMakerOrderService.GetPpcPurchasesForOrderAsync(orderNo);

                // Fetch next batch sequence from legacy logic and increment for each planned line
                string nextBatchStr = await MakerPOService.GetNextBatchNoAsync();
                int batchSeq = 1;
                if (!string.IsNullOrWhiteSpace(nextBatchStr) && nextBatchStr.StartsWith("B17") && int.TryParse(nextBatchStr.Substring(3), out int parsedBatch))
                {
                    batchSeq = parsedBatch;
                }

                // Set default batch & return dates if not set
                foreach (var line in PlannedLines)
                {
                    if (string.IsNullOrWhiteSpace(line.BatchNo) || line.BatchNo == "B171")
                    {
                        line.BatchNo = $"B17{batchSeq++}";
                    }
                    if (line.ReturnDT == default) line.ReturnDT = DateTime.Today.AddDays(30);
                    if (SelectedSteelTypeId.HasValue && !line.SteelType_RefID.HasValue)
                    {
                        line.SteelType_RefID = SelectedSteelTypeId.Value;
                    }
                }

                SelectedOrderLookup = ActiveOrdersWithPurchases.FirstOrDefault(o => o.OrderNo.Equals(orderNo, StringComparison.OrdinalIgnoreCase))
                    ?? OrderHeader;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Order",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void OnOrderSelected(PpcMakerPoOrderHeaderDto? selected)
        {
            if (selected != null && !string.IsNullOrWhiteSpace(selected.OrderNo))
            {
                NavigationManager.NavigateTo($"/production/maker-orders?orderNo={Uri.EscapeDataString(selected.OrderNo)}");
            }
        }

        public async Task GenerateAllPOsAsync()
        {
            var pending = PlannedLines.Where(l => l.IsSelected && !l.IsPosted).ToList();
            if (!pending.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Items Selected",
                    Detail = "There are no pending purchase items selected for PO generation.",
                    Duration = 4000
                });
                return;
            }

            await GeneratePosInternalAsync(pending);
        }

        public async Task GenerateForMakerAsync(int vendId)
        {
            var pending = PlannedLines.Where(l => l.VendID == vendId && l.IsSelected && !l.IsPosted).ToList();
            if (!pending.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Items Selected",
                    Detail = "There are no pending items selected for this maker.",
                    Duration = 4000
                });
                return;
            }

            await GeneratePosInternalAsync(pending);
        }

        private async Task GeneratePosInternalAsync(List<PpcMakerPoItemRowDto> linesToGenerate)
        {
            // Validate that all lines have a maker assigned
            var unassigned = linesToGenerate.FirstOrDefault(l => l.VendID <= 0);
            if (unassigned != null)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Missing Maker Assignment",
                    Detail = $"Item '{unassigned.CompItemID}' has no Maker assigned. Please assign a Maker before generating.",
                    Duration = 5000
                });
                return;
            }

            try
            {
                IsSaving = true;

                var request = new GenerateMakerPoRequest
                {
                    OrderNo = EffectiveOrderNo ?? OrderHeader?.OrderNo ?? string.Empty,
                    IssEmpID = SelectedIssEmp?.EmpID,
                    CountedBy = SelectedCountedBy?.EmpID,
                    SteelType_RefID = SelectedSteelType?.SteelID,
                    Lines = linesToGenerate,
                    UserName = CurrentUserName,
                    UserId = CurrentUserId,
                    MachineName = CurrentMachineName
                };

                var result = await PpcMakerOrderService.GenerateMakerPosAsync(request);

                if (result.Success)
                {
                    var poListStr = string.Join(", ", result.GeneratedMasterPoNumbers);
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Maker PO(s) Created Successfully",
                        Detail = $"Created {result.TotalRowsCreated} line(s) under Master PO(s): {poListStr}",
                        Duration = 8000
                    });

                    // Reload order data
                    if (!string.IsNullOrWhiteSpace(EffectiveOrderNo))
                    {
                        await LoadOrderDataAsync(EffectiveOrderNo);
                        ActiveOrdersWithPurchases = await PpcMakerOrderService.GetActiveOrdersWithPendingPurchasesAsync();
                    }
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "PO Generation Failed",
                        Detail = result.ErrorMessage ?? "An unknown error occurred during PO generation.",
                        Duration = 7000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Generation Error",
                    Detail = ex.Message,
                    Duration = 7000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public async Task PrintMasterPoAsync(string masterPoNo)
        {
            if (string.IsNullOrWhiteSpace(masterPoNo)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "IssList.rpt",
                    SelectionFormula = $"{{VendIssued.MasterPONo}} = '{masterPoNo}'",
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Copy", "'MAKER COPY'" }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Print Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        public void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/production/order-planning");
            }
        }

        // BlazoredTypeahead search helpers
        public Task<IEnumerable<PpcMakerPoOrderHeaderDto>> SearchOrdersAsync(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Task.FromResult<IEnumerable<PpcMakerPoOrderHeaderDto>>(ActiveOrdersWithPurchases.Take(25));
            }

            var filtered = ActiveOrdersWithPurchases
                .Where(o => o.OrderNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                         || o.CustomerName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                         || (o.CustomerPoNo != null && o.CustomerPoNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)))
                .Take(25);

            return Task.FromResult(filtered);
        }

        public Task<IEnumerable<EmployeeLookupModel>> SearchEmployeesAsync(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Task.FromResult<IEnumerable<EmployeeLookupModel>>(Employees.Take(25));
            }

            var filtered = Employees
                .Where(e => e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                         || e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                         || e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .Take(25);

            return Task.FromResult(filtered);
        }

        public Task<IEnumerable<MakerPOLookupModel>> SearchMakersAsync(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Task.FromResult<IEnumerable<MakerPOLookupModel>>(AllMakers.Take(30));
            }

            var filtered = AllMakers
                .Where(m => m.VenderName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                         || m.VendID1.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .Take(30);

            return Task.FromResult(filtered);
        }

        public void OnMakerSelectedForLine(PpcMakerPoItemRowDto line, MakerPOLookupModel? selectedMaker)
        {
            if (selectedMaker != null)
            {
                line.VendID = (int)selectedMaker.VendID;
                line.MakerName = selectedMaker.VenderName;
            }
            else
            {
                line.VendID = 0;
                line.MakerName = "Unassigned";
            }
        }
    }
}
