using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ReportModels;
using DataAccessLibrary.Models.ViewModels.Dashboard;
using Impulse.Services;
using Impulse.Services.Dashboard;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Dashboards.CommandCenter
{
    public partial class CommandCenterDashboard : ComponentBase
    {
        [Inject]
        public ICommandCenterDashboardService DashboardService { get; set; } = default!;

        [Inject]
        public IReportNavigationService ReportNavigationService { get; set; } = default!;

        [Inject]
        public NotificationService NotificationService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        protected CommandCenterDashboardData Data { get; set; } = new();
        protected bool IsLoading { get; set; } = false;
        protected bool IsDataLoaded { get; set; } = false;

        // ── 2-Level Stock Drilldown Modal State ──────────────────────────────
        protected bool IsStockModalOpen { get; set; } = false;
        protected bool IsModalLoading { get; set; } = false;
        protected string ModalStockTitle { get; set; } = "Inventory Breakdown";
        protected int CurrentStockType { get; set; } = 0;
        protected bool IsShowingItemDetail { get; set; } = false;
        protected string SelectedGroupName { get; set; } = string.Empty;
        protected int SelectedGroupID { get; set; } = 0;
        protected string ModalSearchText { get; set; } = string.Empty;

        protected List<StockGroupSummaryItem> GroupSummaries { get; set; } = new();
        protected List<StockGroupItemDetail> ItemDetails { get; set; } = new();

        protected IEnumerable<StockGroupSummaryItem> FilteredGroupSummary =>
            string.IsNullOrWhiteSpace(ModalSearchText)
                ? GroupSummaries
                : GroupSummaries.Where(x => x.GroupName.Contains(ModalSearchText, StringComparison.OrdinalIgnoreCase));

        protected IEnumerable<StockGroupItemDetail> FilteredItemDetails =>
            string.IsNullOrWhiteSpace(ModalSearchText)
                ? ItemDetails
                : ItemDetails.Where(x =>
                    x.ItemID.Contains(ModalSearchText, StringComparison.OrdinalIgnoreCase) ||
                    x.ItemName.Contains(ModalSearchText, StringComparison.OrdinalIgnoreCase));

        protected override async Task OnInitializedAsync()
        {
            await RefreshDataAsync();
        }

        protected async Task RefreshDataAsync()
        {
            IsLoading = true;
            try
            {
                Data = await DashboardService.GetCommandCenterDataAsync();
                IsDataLoaded = true;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Command Center Sync Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ── Inventory Modal Drilldowns ───────────────────────────────────────
        protected async Task OpenStockGroupModal(int stockType, string title)
        {
            CurrentStockType = stockType;
            ModalStockTitle = title;
            IsShowingItemDetail = false;
            SelectedGroupName = string.Empty;
            SelectedGroupID = 0;
            ModalSearchText = string.Empty;
            IsStockModalOpen = true;
            IsModalLoading = true;

            try
            {
                GroupSummaries = await DashboardService.GetStockGroupSummaryAsync(stockType);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Stock Groups",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
            finally
            {
                IsModalLoading = false;
            }
        }

        protected async Task DrillDownIntoGroup(StockGroupSummaryItem group)
        {
            SelectedGroupID = group.GroupID;
            SelectedGroupName = group.GroupName;
            IsShowingItemDetail = true;
            ModalSearchText = string.Empty;
            IsModalLoading = true;

            try
            {
                ItemDetails = await DashboardService.GetStockGroupDetailAsync(CurrentStockType, group.GroupID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Item Details",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
            finally
            {
                IsModalLoading = false;
            }
        }

        protected void BackToGroupSummary()
        {
            IsShowingItemDetail = false;
            ModalSearchText = string.Empty;
        }

        protected void CloseStockModal()
        {
            IsStockModalOpen = false;
            IsShowingItemDetail = false;
            GroupSummaries.Clear();
            ItemDetails.Clear();
        }

        // ── Reports Integration ──────────────────────────────────────────────
        protected async Task PrintCashBookReport()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptCashBook",
                    SelectionFormula = $"{{VBankCashPaymentReceipt.VDate}} = #{DateTime.Today:yyyy-MM-dd}#"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }

        protected async Task PrintUnshippedOrderList()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptFOrderList",
                    SelectionFormula = "{VFOrderList.TotalShippedQty} < {VFOrderList.TotalOrderQty} AND ISNULL({VFOrderList.Cancelled}, 0) = 0 AND {VFOrderList.CustCode} <> 'Stock'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }

        protected async Task PrintAbsentSheet()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptEmpAbsentSheetSummary",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@DTFrom", DateTime.Today },
                        { "@DTTo", DateTime.Today }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }

        protected async Task PrintLoanBalance()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptEmpLoanBalance"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }

        protected async Task PrintTotalExportCurrentFY()
        {
            try
            {
                DateTime today = DateTime.Today;
                DateTime fyFrom = today.Month >= 7 ? new DateTime(today.Year, 7, 1) : new DateTime(today.Year - 1, 7, 1);
                DateTime fyTo = today.Month >= 7 ? new DateTime(today.Year + 1, 6, 30) : new DateTime(today.Year, 6, 30);

                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "TotalExport",
                    SelectionFormula = $"{{VTotalExport.DT}} >= #{fyFrom:yyyy-MM-dd}# AND {{VTotalExport.DT}} <= #{fyTo:yyyy-MM-dd}#"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }

        protected async Task PrintRunningLots(long processId = 0)
        {
            try
            {
                string formula = processId > 0 ? $"{{VRunningLots.ProcessID}} = {processId}" : string.Empty;
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "RunningLots",
                    SelectionFormula = formula
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }

        protected async Task PrintReWorkLots()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "ReWorkLots",
                    SelectionFormula = "{VVendRcvdDetailReWorkDetail.Qty} > {VVendRcvdDetailReWorkDetail.IssQty}"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }

        protected async Task PrintOpenMakerOrders()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "PendingPurchaseOrdersList"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }

        protected async Task PrintForgingOpenOrders()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "RM_PO_List",
                    Parameters = new Dictionary<string, object>
                    {
                        { "Heading", "Forging Open Orders" }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 3500
                });
            }
        }
    }
}
