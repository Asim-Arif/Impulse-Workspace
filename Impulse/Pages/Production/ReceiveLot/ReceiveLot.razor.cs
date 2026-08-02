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

namespace Impulse.Pages.Production.ReceiveLot
{
    public partial class ReceiveLot
    {
        [Inject]
        public IReceiveLotService ReceiveLotService { get; set; } = default!;

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
        public string MillCertNo { get; set; } = string.Empty;

        public LotSearchResultModel? LotHeader { get; set; }
        public List<MasterPOLineItemModel> Lines { get; set; } = new List<MasterPOLineItemModel>();

        public List<EmployeeLookupModel> AvailableEmployees { get; set; } = new List<EmployeeLookupModel>();
        public EmployeeLookupModel? SelectedWorker { get; set; }
        public List<EmployeeLookupModel> SelectedWorkers { get; set; } = new List<EmployeeLookupModel>();

        public bool PrintSlip { get; set; } = true;
        public bool IsSearching { get; set; } = false;
        public bool IsSaving { get; set; } = false;

        public bool ShowImagePreviewModal { get; set; } = false;
        public string PreviewImageBase64 { get; set; } = string.Empty;
        public string PreviewItemCode { get; set; } = string.Empty;

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

        public decimal TotalIssuedQty => Lines.Sum(l => l.IssQty);
        public decimal TotalPrevRcvdQty => Lines.Sum(l => l.PrevRcvdQty);
        public decimal TotalPendingQty => Lines.Sum(l => l.PendingQty);
        public decimal TotalReceivingQty => Lines.Sum(l => l.ReceivingQty);
        public decimal TotalReceivingValue => Lines.Sum(l => l.ReceivingQty * l.Rate);

        public async Task HandleSearchKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await SearchLotAsync();
            }
        }

        public async Task SearchLotAsync()
        {
            if (string.IsNullOrWhiteSpace(SearchLotNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Input Required",
                    Detail = "Please enter a valid Lot Number.",
                    Duration = 4000
                });
                return;
            }

            IsSearching = true;
            try
            {
                var searchResult = await ReceiveLotService.SearchLotAsync(SearchLotNo.Trim());
                if (searchResult == null)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Lot Not Found",
                        Detail = $"No issued Lot found with Lot No: {SearchLotNo}",
                        Duration = 5000
                    });
                    LotHeader = null;
                    Lines.Clear();
                    return;
                }

                if (!searchResult.Authorized)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Order Not Authorized",
                        Detail = $"Lot [{searchResult.LotNo}] has not been authorized yet.",
                        Duration = 5000
                    });
                    LotHeader = null;
                    Lines.Clear();
                    return;
                }

                if (searchResult.AlreadyReceived)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Already Received",
                        Detail = $"Lot No [{searchResult.LotNo}] has already been received at this process step.",
                        Duration = 5000
                    });
                    LotHeader = null;
                    Lines.Clear();
                    return;
                }

                LotHeader = searchResult;
                Lines = await ReceiveLotService.GetLotLinesAsync(LotHeader.LotNo, LotHeader.EntryID);

                // Auto-fill ReceivingQty to PendingQty by default
                foreach (var line in Lines)
                {
                    line.ReceivingQty = Math.Max(0, line.PendingQty);
                }

                if (LotHeader.IsFactoryMaker)
                {
                    AvailableEmployees = await MakerPOService.GetEmployeesAsync();
                }

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Lot Loaded",
                    Detail = $"Lot [{LotHeader.LotNo}] loaded for Process: {LotHeader.ProcessName}, Vendor: {LotHeader.VendorName}",
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
            finally
            {
                IsSearching = false;
            }
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

        public void AddWorker()
        {
            if (SelectedWorker != null && !SelectedWorkers.Any(w => w.EmpID == SelectedWorker.EmpID))
            {
                SelectedWorkers.Add(SelectedWorker);
                SelectedWorker = null;
            }
        }

        public void RemoveWorker(EmployeeLookupModel emp)
        {
            SelectedWorkers.Remove(emp);
        }

        public void ClearForm()
        {
            SearchLotNo = string.Empty;
            MillCertNo = string.Empty;
            LotHeader = null;
            Lines.Clear();
            SelectedWorkers.Clear();
            SelectedWorker = null;
        }

        public async Task SaveReceiving()
        {
            if (LotHeader == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Error",
                    Detail = "No Lot loaded.",
                    Duration = 4000
                });
                return;
            }

            if (LotHeader.IsFactoryMaker && !SelectedWorkers.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Workers Required",
                    Detail = "Please select at least 1 in-house factory worker for this lot receiving.",
                    Duration = 4000
                });
                return;
            }

            var activeReceivingLines = Lines.Where(l => l.ReceivingQty > 0).ToList();
            if (!activeReceivingLines.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Invalid Quantities",
                    Detail = "Please enter receiving quantity greater than zero for at least one item line.",
                    Duration = 4000
                });
                return;
            }

            foreach (var line in activeReceivingLines)
            {
                if (line.ReceivingQty > line.PendingQty)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Quantity Exceeds Pending Limit",
                        Detail = $"Item [{line.ItemCode}] receiving quantity ({line.ReceivingQty:N0}) exceeds pending quantity ({line.PendingQty:N0}).",
                        Duration = 5000
                    });
                    return;
                }
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

                var headerPayload = new CreatePOReceivingHeaderModel
                {
                    IssuanceRefID = LotHeader.EntryID,
                    VendID = LotHeader.VendID,
                    ProcessID = LotHeader.ProcessID,
                    DT = DateTime.Today,
                    MillCertNo = MillCertNo,
                    FactoryEmpIDs = SelectedWorkers.Select(w => w.EmpID).ToList()
                };

                var linePayloads = activeReceivingLines.Select(l => new CreatePOReceivingLineModel
                {
                    VendIssdDetailEntryID = l.VendIssdDetailEntryID,
                    ItemCode = l.ItemCode,
                    OrderNo = l.OrderNo,
                    Rate = l.Rate,
                    RcvdQty = l.ReceivingQty,
                    CountedBy = l.CountedBy,
                    LotNo = string.IsNullOrWhiteSpace(l.LotNo) ? LotHeader.LotNo : l.LotNo,
                    ReWorkLot = l.ReWorkLot,
                    RepairType = l.RepairType,
                    ProcessID = LotHeader.ProcessID
                }).ToList();

                long rcvHeaderId = await ReceiveLotService.SaveLotReceivingAsync(headerPayload, linePayloads, userName, userId, machineName);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Lot Receiving Saved",
                    Detail = $"Receiving posted successfully for Lot [{LotHeader.LotNo}]. Entry ID: {rcvHeaderId}",
                    Duration = 5000
                });

                if (PrintSlip)
                {
                    await ReportNavigationService.PrintReportAsync(new ReportRequest
                    {
                        ReportName = "RcvSlip.rpt",
                        SelectionFormula = $"{{VendReceived.EntryID}}={rcvHeaderId}"
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
