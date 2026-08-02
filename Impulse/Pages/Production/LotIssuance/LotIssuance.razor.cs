using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Components.Authorization;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Impulse.Services;

namespace Impulse.Pages.Production.LotIssuance
{
    public partial class LotIssuance
    {
        [Inject]
        public ILotIssuanceService LotIssuanceService { get; set; } = default!;

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

        public string SearchLotNo { get; set; } = string.Empty;
        public int SelectedProcessID { get; set; } = 0;
        public string BatchNo { get; set; } = string.Empty;
        public List<ProcessPOLookupModel> Processes { get; set; } = new List<ProcessPOLookupModel>();

        public MakerPOLookupModel? SelectedMaker { get; set; }
        public List<MakerPOLookupModel> AvailableMakers { get; set; } = new List<MakerPOLookupModel>();

        public bool SampleProvided { get; set; } = false;
        public bool DrawingProvided { get; set; } = false;
        public bool ForgingProvided { get; set; } = false;
        public bool SteelProvided { get; set; } = false;

        public List<AvailableLotIssuanceItemModel> StagedItems { get; set; } = new List<AvailableLotIssuanceItemModel>();

        public List<EmployeeLookupModel> AvailableEmployees { get; set; } = new List<EmployeeLookupModel>();
        public EmployeeLookupModel? SelectedWorker { get; set; }
        public EmployeeLookupModel? SelectedCountedBy { get; set; }

        public bool PrintSlip { get; set; } = true;
        public bool IsSearching { get; set; } = false;
        public bool IsSaving { get; set; } = false;

        public bool IsFactoryMaker => SelectedMaker != null &&
            (SelectedMaker.VenderName.Contains("FACTORY", StringComparison.OrdinalIgnoreCase) ||
             SelectedMaker.VendID1.Contains("FAC", StringComparison.OrdinalIgnoreCase) ||
             SelectedMaker.VendID == 79);

        public bool ShowImagePreviewModal { get; set; } = false;
        public string PreviewImageBase64 { get; set; } = string.Empty;
        public string PreviewItemCode { get; set; } = string.Empty;

        public decimal TotalAvailableQty => StagedItems.Sum(i => i.AvailableQty);
        public decimal TotalIssuanceQty => StagedItems.Sum(i => i.IssuanceQty);
        public decimal TotalIssuanceValue => StagedItems.Sum(i => i.IssuanceQty * i.Rate);

        protected override async Task OnInitializedAsync()
        {
            Processes = await LotIssuanceService.GetProcessesAsync();
            AvailableEmployees = await MakerPOService.GetEmployeesAsync();
        }

        public async Task OnProcessChangedAsync(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int processId))
            {
                SelectedProcessID = processId;
                SelectedMaker = null;
                AvailableMakers.Clear();

                if (SelectedProcessID > 0)
                {
                    AvailableMakers = await LotIssuanceService.GetMakersForProcessAsync(SelectedProcessID);
                }
            }
        }

        public Task<IEnumerable<MakerPOLookupModel>> SearchMakers(string searchText)
        {
            if (AvailableMakers == null || !AvailableMakers.Any())
                return Task.FromResult(Enumerable.Empty<MakerPOLookupModel>());

            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<MakerPOLookupModel>>(AvailableMakers);

            return Task.FromResult<IEnumerable<MakerPOLookupModel>>(
                AvailableMakers.Where(m => (m.VenderName != null && m.VenderName.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                                         || (m.VendID1 != null && m.VendID1.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
        }

        public Task<IEnumerable<EmployeeLookupModel>> SearchEmployees(string searchText)
        {
            if (AvailableEmployees == null || !AvailableEmployees.Any())
                return Task.FromResult(Enumerable.Empty<EmployeeLookupModel>());

            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<EmployeeLookupModel>>(AvailableEmployees);

            return Task.FromResult<IEnumerable<EmployeeLookupModel>>(
                AvailableEmployees.Where(e => (e.Name != null && e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                                           || (e.EmpID != null && e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                                           || (e.Designation != null && e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
        }

        public async Task HandleLotSearchKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await LookupLotAsync();
            }
        }

        public async Task LookupLotAsync()
        {
            if (string.IsNullOrWhiteSpace(SearchLotNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Input Required",
                    Detail = "Please enter a valid Lot Number to search.",
                    Duration = 4000
                });
                return;
            }

            IsSearching = true;
            try
            {
                var result = await LotIssuanceService.LookupLotForIssuanceAsync(SearchLotNo.Trim());

                if (result.AlreadyIssued)
                {
                    StagedItems.Clear();
                    SelectedProcessID = 0;
                    BatchNo = string.Empty;
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Already Issued",
                        Detail = result.Message,
                        Duration = 5000
                    });
                    return;
                }

                if (!result.IsFound || !result.Items.Any())
                {
                    StagedItems.Clear();
                    SelectedProcessID = 0;
                    BatchNo = string.Empty;
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "No Issuable Items Found",
                        Detail = result.Message,
                        Duration = 5000
                    });
                    return;
                }

                var lotItems = result.Items;
                var firstItem = lotItems.First();
                if (firstItem.TargetProcessID > 0)
                {
                    SelectedProcessID = firstItem.TargetProcessID;
                    AvailableMakers = await LotIssuanceService.GetMakersForProcessAsync(SelectedProcessID);
                }

                if (!string.IsNullOrWhiteSpace(firstItem.BatchNo))
                {
                    BatchNo = firstItem.BatchNo;
                }

                foreach (var item in lotItems)
                {
                    if (!StagedItems.Any(s => s.VendIssdDetailEntryID == item.VendIssdDetailEntryID))
                    {
                        item.IssuanceQty = item.AvailableQty;
                        StagedItems.Add(item);
                    }
                }

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Lot Details Loaded",
                    Detail = $"Loaded {lotItems.Count} item(s) for Lot [{SearchLotNo}] at Process: {firstItem.TargetProcessName}.",
                    Duration = 4000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Search Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSearching = false;
            }
        }

        public void RemoveStagedItem(AvailableLotIssuanceItemModel item)
        {
            StagedItems.Remove(item);
        }

        public void OpenImagePreview(string imageBase64, string itemCode)
        {
            PreviewImageBase64 = imageBase64;
            PreviewItemCode = itemCode;
            ShowImagePreviewModal = true;
        }

        public void CloseImagePreview()
        {
            ShowImagePreviewModal = false;
            PreviewImageBase64 = string.Empty;
            PreviewItemCode = string.Empty;
        }

        public void ClearForm()
        {
            SearchLotNo = string.Empty;
            SelectedProcessID = 0;
            BatchNo = string.Empty;
            SelectedMaker = null;
            SelectedWorker = null;
            SelectedCountedBy = null;
            SampleProvided = false;
            DrawingProvided = false;
            ForgingProvided = false;
            SteelProvided = false;
            StagedItems.Clear();
        }

        public async Task SaveLotIssuance()
        {
            if (SelectedProcessID <= 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Error",
                    Detail = "Please enter a valid Lot No to fetch the process.",
                    Duration = 4000
                });
                return;
            }

            if (SelectedMaker == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Error",
                    Detail = "Please select a target Maker/Vendor.",
                    Duration = 4000
                });
                return;
            }

            if (IsFactoryMaker && SelectedWorker == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Worker Required",
                    Detail = "Please select an in-house factory worker.",
                    Duration = 4000
                });
                return;
            }

            var validLines = StagedItems.Where(i => i.IssuanceQty > 0).ToList();
            if (!validLines.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Invalid Quantities",
                    Detail = "No staged items available for issuance.",
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

                var headerPayload = new CreateLotIssuanceHeaderModel
                {
                    VendID = SelectedMaker.VendID,
                    ProcessID = SelectedProcessID,
                    DT = DateTime.Today,
                    SampleProvided = SampleProvided,
                    DrawingProvided = DrawingProvided,
                    ForgingProvided = ForgingProvided,
                    SteelProvided = SteelProvided,
                    IssEmpID = SelectedWorker?.EmpID ?? "",
                    CountedBy = SelectedCountedBy?.Name ?? ""
                };

                var linePayloads = validLines.Select(i => new CreateLotIssuanceLineModel
                {
                    IssueRefID = i.VendIssdDetailEntryID,
                    ItemCode = i.ItemCode,
                    OrderNo = i.OrderNo,
                    Rate = i.Rate,
                    IssQty = i.IssuanceQty,
                    LotNo = i.LotNo,
                    BatchNo = BatchNo,
                    RcvProcessID = i.ReturnProcessID > 0 ? i.ReturnProcessID : SelectedProcessID,
                    ReturnDT = DateTime.Today,
                    Priority = i.Priority,
                    ReWorkLot = i.ReWorkLot,
                    RepairType = i.RepairType,
                    CountedBy = SelectedCountedBy?.Name ?? ""
                }).ToList();

                long headerId = await LotIssuanceService.SaveLotIssuanceAsync(headerPayload, linePayloads, userName, userId, machineName);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Lot Issuance Posted",
                    Detail = $"Lot Issuance saved successfully. Entry ID: {headerId}",
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
