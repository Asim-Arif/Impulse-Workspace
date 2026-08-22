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
    public class FinancialAssetChartPoint
    {
        public string Category { get; set; } = string.Empty;
        public double Value { get; set; }
        public string Color { get; set; } = string.Empty;
    }

    public class ExportTurnoverChartPoint
    {
        public string Category { get; set; } = string.Empty;
        public double ValueInMillions { get; set; }
        public string FormattedValue { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
    }

    public class InventoryValuationChartPoint
    {
        public string TierName { get; set; } = string.Empty;
        public double ValueInMillions { get; set; }
        public long Quantity { get; set; }
        public string FormattedValue { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public int StockType { get; set; }
    }

    public class ManufacturingStationPoint
    {
        public string StationName { get; set; } = string.Empty;
        public int LotCount { get; set; }
        public long Quantity { get; set; }
        public double PercentageOfWip { get; set; }
        public long ProcessID { get; set; }
        public string Color { get; set; } = string.Empty;
        public string Icon { get; set; } = string.Empty;
    }

    public partial class CommandCenterVisualDashboard : ComponentBase
    {
        [Inject] public ICommandCenterDashboardService DashboardService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;

        public CommandCenterDashboardData Data { get; set; } = new();
        public bool IsLoading { get; set; } = false;
        public bool IsDataLoaded { get; set; } = false;

        // View Mode: 0 = Executive Analytics Cockpit, 1 = Detailed Operations
        public int ActiveViewMode { get; set; } = 0;

        // Chart Projections
        public int SelectedTrendMetric { get; set; } = 0; // 0 = Revenue (PKR), 1 = Volume (Units)
        public int SelectedBarMetric { get; set; } = 0; // 0 = Volume (Units), 1 = Invoices Count
        public List<FinancialAssetChartPoint> FinancialAssetChartData { get; set; } = new();
        public List<ExportTurnoverChartPoint> ExportTurnoverChartData { get; set; } = new();
        public List<InventoryValuationChartPoint> InventoryValuationChartData { get; set; } = new();
        public List<ManufacturingStationPoint> BottleneckStationsData { get; set; } = new();

        // ── 2-Level Stock Drilldown Modal State ──────────────────────────────
        public bool IsStockModalOpen { get; set; } = false;
        public bool IsModalLoading { get; set; } = false;
        public string ModalStockTitle { get; set; } = "Inventory Valuation Breakdown";
        public int CurrentStockType { get; set; } = 0;
        public bool IsShowingItemDetail { get; set; } = false;
        public string SelectedGroupName { get; set; } = string.Empty;
        public int SelectedGroupID { get; set; } = 0;
        public string ModalSearchText { get; set; } = string.Empty;

        public List<StockGroupSummaryItem> GroupSummaries { get; set; } = new();
        public List<StockGroupItemDetail> ItemDetails { get; set; } = new();

        public IEnumerable<StockGroupSummaryItem> FilteredGroupSummary =>
            string.IsNullOrWhiteSpace(ModalSearchText)
                ? GroupSummaries
                : GroupSummaries.Where(x => x.GroupName.Contains(ModalSearchText, StringComparison.OrdinalIgnoreCase));

        public IEnumerable<StockGroupItemDetail> FilteredItemDetails =>
            string.IsNullOrWhiteSpace(ModalSearchText)
                ? ItemDetails
                : ItemDetails.Where(x =>
                    x.ItemID.Contains(ModalSearchText, StringComparison.OrdinalIgnoreCase) ||
                    x.ItemName.Contains(ModalSearchText, StringComparison.OrdinalIgnoreCase));

        protected override async Task OnInitializedAsync()
        {
            await RefreshDataAsync();
        }

        public async Task RefreshDataAsync()
        {
            IsLoading = true;
            try
            {
                Data = await DashboardService.GetCommandCenterDataAsync();
                BuildVisualProjections();
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

        private void BuildVisualProjections()
        {
            // 1. Financial Assets Allocation Donut
            FinancialAssetChartData = new List<FinancialAssetChartPoint>
            {
                new() { Category = "Cash-in-Hand", Value = (double)Data.Financial.CashInHand, Color = "#107C41" },
                new() { Category = "Bank Balances", Value = (double)Data.Financial.BankBalance, Color = "#118DFF" },
                new() { Category = "Accounts Receivable", Value = (double)Data.Financial.AccountsReceivable, Color = "#EAA300" },
                new() { Category = "Fixed Assets", Value = (double)Data.Financial.FixedAssetsCost, Color = "#7B1FA2" }
            };

            // 2. Export Turnover Horizon
            ExportTurnoverChartData = new List<ExportTurnoverChartPoint>
            {
                new() { Category = "In-Hand Orders", ValueInMillions = Math.Round((double)Data.Export.InHandValue / 1_000_000.0, 2), FormattedValue = $"PKR {(Data.Export.InHandValue / 1_000_000m):N2}M", Color = "#118DFF" },
                new() { Category = "Current FY Sales", ValueInMillions = Math.Round((double)Data.Export.TotalSalesCurrentFY / 1_000_000.0, 2), FormattedValue = $"PKR {(Data.Export.TotalSalesCurrentFY / 1_000_000m):N2}M", Color = "#00BFA5" },
                new() { Category = "Lifetime Sales", ValueInMillions = Math.Round((double)Data.Export.LifetimeTotalSales / 1_000_000.0, 2), FormattedValue = $"PKR {(Data.Export.LifetimeTotalSales / 1_000_000m):N2}M", Color = "#107C41" }
            };

            // 3. Multi-Tier Inventory Valuation Donut
            InventoryValuationChartData = new List<InventoryValuationChartPoint>
            {
                new() { TierName = "Finished Goods", ValueInMillions = (double)Data.Stock.FinishedStockValue, Quantity = Data.Stock.FinishedStockQty, FormattedValue = $"PKR {Data.Stock.FinishedStockValue:N0}", Color = "#118DFF", StockType = 0 },
                new() { TierName = "Semi-Finished", ValueInMillions = (double)Data.Stock.SFStockValue, Quantity = Data.Stock.SFStockQty, FormattedValue = $"PKR {Data.Stock.SFStockValue:N0}", Color = "#00BFA5", StockType = 1 },
                new() { TierName = "Forgings Stock", ValueInMillions = (double)Data.Stock.ForgingStockValue, Quantity = Data.Stock.ForgingStockQty, FormattedValue = $"PKR {Data.Stock.ForgingStockValue:N0}", Color = "#E65100", StockType = 2 },
                new() { TierName = "Other Materials", ValueInMillions = (double)Data.Stock.OtherStockValue, Quantity = Data.Stock.OtherStockQty, FormattedValue = $"PKR {Data.Stock.OtherStockValue:N0}", Color = "#7B1FA2", StockType = 3 }
            };

            // 4. Shop Floor Bottleneck Stations
            double totalWip = Data.Production.RunningLotsQty > 0 ? Data.Production.RunningLotsQty : 1;
            BottleneckStationsData = new List<ManufacturingStationPoint>
            {
                new() { StationName = "Tempering (Heat Treat)", LotCount = Data.Production.LotsAtTemperCount, Quantity = Data.Production.LotsAtTemperQty, PercentageOfWip = Math.Round(((double)Data.Production.LotsAtTemperQty / totalWip) * 100.0, 1), ProcessID = Data.Production.HeatTreatmentProcessID, Color = "#E81123", Icon = "fas fa-fire-flame-curved" },
                new() { StationName = "Polishing Department", LotCount = Data.Production.LotsAtPolishingCount, Quantity = Data.Production.LotsAtPolishingQty, PercentageOfWip = Math.Round(((double)Data.Production.LotsAtPolishingQty / totalWip) * 100.0, 1), ProcessID = Data.Production.PolishingProcessID, Color = "#118DFF", Icon = "fas fa-wand-magic-sparkles" },
                new() { StationName = "QC 100% Inspection", LotCount = Data.Production.LotsAtQCCount, Quantity = Data.Production.LotsAtQCQty, PercentageOfWip = Math.Round(((double)Data.Production.LotsAtQCQty / totalWip) * 100.0, 1), ProcessID = Data.Production.QCProcessID, Color = "#107C41", Icon = "fas fa-clipboard-check" },
                new() { StationName = "Final Packaging", LotCount = Data.Production.LotsAtPackedCount, Quantity = Data.Production.LotsAtPackedQty, PercentageOfWip = Math.Round(((double)Data.Production.LotsAtPackedQty / totalWip) * 100.0, 1), ProcessID = Data.Production.PackingProcessID, Color = "#7B1FA2", Icon = "fas fa-box" }
            };
        }

        // ── 2-Level Stock Drilldown Modal ────────────────────────────────────
        public async Task OpenStockGroupModal(int stockType, string title)
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

        public async Task DrillDownIntoGroup(StockGroupSummaryItem group)
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

        public void BackToGroupSummary()
        {
            IsShowingItemDetail = false;
            ModalSearchText = string.Empty;
        }

        public void CloseStockModal()
        {
            IsStockModalOpen = false;
            IsShowingItemDetail = false;
            GroupSummaries.Clear();
            ItemDetails.Clear();
        }

        // ── Report Integrations ──────────────────────────────────────────────
        public async Task PrintCashBookReport()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "CashBookReport.rpt",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@DTFrom", DateTime.Today },
                        { "@DTTo", DateTime.Today }
                    },
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "FromTo", $"'{DateTime.Today:dd-MMM-yyyy} to {DateTime.Today:dd-MMM-yyyy}'" }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }

        public async Task PrintUnshippedOrderList()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptFOrderList.rpt",
                    SelectionFormula = "{VFOrderList.TotalShippedQty} < {VFOrderList.TotalOrderQty} AND (IsNull({VFOrderList.Cancelled}) OR {VFOrderList.Cancelled} = 0) AND {VFOrderList.CustCode} <> 'Stock'",
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "ForCustomer", "'<All Customers>'" },
                        { "ForFromTo", "'All Pending'" }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }

        public async Task PrintAbsentSheet()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Payroll/rptEmpAbsentSheetSummary.rpt",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@DTFrom", DateTime.Today },
                        { "@DTTo", DateTime.Today }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }

        public async Task PrintLoanBalance()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Payroll/rptEmpLoanBalance.rpt",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@DT", DateTime.Today },
                        { "@InActive", false }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }

        public async Task PrintTotalExportCurrentFY()
        {
            try
            {
                DateTime today = DateTime.Today;
                DateTime fyFrom = today.Month >= 7 ? new DateTime(today.Year, 7, 1) : new DateTime(today.Year - 1, 7, 1);
                DateTime fyTo = today.Month >= 7 ? new DateTime(today.Year + 1, 6, 30) : new DateTime(today.Year, 6, 30);

                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "TotalExport.rpt",
                    SelectionFormula = $"{{VTotalExport.DT}} in Date({fyFrom.Year}, {fyFrom.Month}, {fyFrom.Day}) to Date({fyTo.Year}, {fyTo.Month}, {fyTo.Day})",
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "CustomAmt", "false" },
                        { "SubHeading", $"'{fyFrom:dd-MMM-yyyy} to {fyTo:dd-MMM-yyyy}'" }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }

        public async Task PrintRunningLots(long processId = 0)
        {
            try
            {
                string formula = processId > 0 ? $"{{VRunningLots.ProcessID}} = {processId}" : string.Empty;
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "RunningLots.rpt",
                    SelectionFormula = formula
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }

        public async Task PrintReWorkLots()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "ReWorkLots.rpt",
                    SelectionFormula = "{VVendRcvdDetailReWorkDetail.Qty} > {VVendRcvdDetailReWorkDetail.IssQty}"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }

        public async Task PrintOpenMakerOrders()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Maker_Open_POs.rpt"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }

        public async Task PrintForgingOpenOrders()
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "RM_PO_List.rpt",
                    SelectionFormula = "{VVendOrdersToRcv.OrderNo} = {VVendOrders.OrderNo} AND {@BalanceQty} > 0 AND {Material.ForgingGroupID} <> 0",
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Heading", "'Forging Open Orders'" }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Report Error", Detail = ex.Message, Duration = 3500 });
            }
        }
    }
}
