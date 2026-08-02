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

namespace Impulse.Pages.Production.ReceiveAgainstPO
{
    public partial class ReceiveAgainstPO
    {
        [Inject]
        public IReceiveAgainstPOService ReceiveService { get; set; } = default!;

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

        public string SearchReceiptNo { get; set; } = string.Empty;
        public string MillCertNo { get; set; } = string.Empty;

        public MasterPOSearchResultModel? MasterPO { get; set; }
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
                await SearchMasterPOAsync();
            }
        }

        public async Task SearchMasterPOAsync()
        {
            if (string.IsNullOrWhiteSpace(SearchReceiptNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Input Required",
                    Detail = "Please enter a valid Master PO Receipt ID.",
                    Duration = 4000
                });
                return;
            }

            IsSearching = true;
            try
            {
                var searchResult = await ReceiveService.SearchMasterPOAsync(SearchReceiptNo.Trim());
                if (searchResult == null)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Master PO Not Found",
                        Detail = $"No Master PO found with Receipt ID: {SearchReceiptNo}",
                        Duration = 5000
                    });
                    MasterPO = null;
                    Lines.Clear();
                    return;
                }

                if (!searchResult.Authorized)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Order Not Authorized",
                        Detail = $"Master PO [{searchResult.RecieptID}] has not been authorized yet.",
                        Duration = 5000
                    });
                    MasterPO = null;
                    Lines.Clear();
                    return;
                }

                if (searchResult.Closed)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Master PO Closed",
                        Detail = $"Master PO [{searchResult.RecieptID}] is closed and cannot receive items.",
                        Duration = 5000
                    });
                    MasterPO = null;
                    Lines.Clear();
                    return;
                }

                MasterPO = searchResult;
                Lines = await ReceiveService.GetMasterPOLinesAsync(MasterPO.EntryID);

                // Auto-fill ReceivingQty to PendingQty by default for user convenience
                foreach (var line in Lines)
                {
                    line.ReceivingQty = Math.Max(0, line.PendingQty);
                }

                if (MasterPO.IsFactoryMaker)
                {
                    AvailableEmployees = await MakerPOService.GetEmployeesAsync();
                }

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Master PO Loaded",
                    Detail = $"PO [{MasterPO.RecieptID}] loaded for Process: {MasterPO.ProcessName}, Vendor: {MasterPO.VendorName}",
                    Duration = 4000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Searching Master PO",
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
            SearchReceiptNo = string.Empty;
            MasterPO = null;
            Lines.Clear();
            SelectedWorkers.Clear();
            SelectedWorker = null;
        }

        public async Task SaveReceiving()
        {
            if (MasterPO == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Error",
                    Detail = "No Master PO loaded.",
                    Duration = 4000
                });
                return;
            }

            if (MasterPO.IsFactoryMaker && !SelectedWorkers.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Workers Required",
                    Detail = "Please select at least 1 in-house factory worker for this receiving.",
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
                    IssuanceRefID = MasterPO.EntryID,
                    VendID = MasterPO.VendID,
                    ProcessID = MasterPO.ProcessID,
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
                    LotNo = l.LotNo,
                    ReWorkLot = l.ReWorkLot,
                    RepairType = l.RepairType,
                    ProcessID = MasterPO.ProcessID
                }).ToList();

                long rcvHeaderId = await ReceiveService.SavePOReceivingAsync(headerPayload, linePayloads, userName, userId, machineName);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Receiving Saved",
                    Detail = $"Receiving posted successfully against PO [{MasterPO.RecieptID}]. Entry ID: {rcvHeaderId}",
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
