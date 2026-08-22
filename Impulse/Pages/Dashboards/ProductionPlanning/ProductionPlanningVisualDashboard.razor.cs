using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Dashboard;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Dashboard;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Dashboards.ProductionPlanning
{
    public class ChartDataPoint
    {
        public string Category { get; set; } = string.Empty;
        public double Value { get; set; }
        public string SubText { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
    }

    public class StageWipChartPoint
    {
        public string StageName { get; set; } = string.Empty;
        public double Quantity { get; set; }
        public int LotCount { get; set; }
        public long ProcessID { get; set; }
        public double PercentageOfTotal { get; set; }
    }

    public class MakerLoadChartPoint
    {
        public string MakerName { get; set; } = string.Empty;
        public double Quantity { get; set; }
        public int PoCount { get; set; }
        public bool HasOverdue { get; set; }
        public int OverdueCount { get; set; }
        public double PercentageOfTotal { get; set; }
    }

    public class DeliveryTimelineChartPoint
    {
        public string TimelineBucket { get; set; } = string.Empty;
        public string BadgeClass { get; set; } = string.Empty;
        public string IconClass { get; set; } = string.Empty;
        public double Quantity { get; set; }
        public int OrderCount { get; set; }
        public double PercentageOfDemand { get; set; }
    }

    public partial class ProductionPlanningVisualDashboard : ComponentBase
    {
        [Inject] public IProductionPlanningDashboardService DashboardService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;
        [Inject] public DialogService DialogService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;

        // Slicer Filters
        public List<LookupItemString> Customers { get; set; } = new();
        public string SelectedCustCode { get; set; } = "0";

        public List<string> Countries { get; set; } = new() { "<Comp>" };
        public string SelectedCountry { get; set; } = "<Comp>";

        public LookupItemString? SelectedArticle { get; set; }
        public string? FilterOrderNo { get; set; }
        public string? SelectedStageFilter { get; set; }

        public ProductionPlanningDashboardData? DashboardData { get; set; }
        public bool IsLoading { get; set; } = false;

        // View Mode: 0 = Executive Visuals, 1 = Split View (Charts + Matrix), 2 = Data Matrix Only
        public int ActiveViewMode { get; set; } = 1;

        // Matrix Tab: 0 = Orders, 1 = WIP Lots, 2 = Maker POs, 3 = Inventory, 4 = Raw Material / Forgings, 5 = Routing & Rates, 6 = Quality / Complaints
        public int ActiveMatrixTab { get; set; } = 0;
        public int InventorySubTab { get; set; } = 0;

        // Cross-Demand Modal
        public bool ShowCrossDemandModal { get; set; } = false;
        public string ActiveForgingCode { get; set; } = string.Empty;
        public List<ForgingCrossDemandItem> CrossDemandItems { get; set; } = new();

        // Visual Calculations & Metrics
        public double DemandCoveragePercent { get; set; } = 0;
        public List<StageWipChartPoint> StageWipChartData { get; set; } = new();
        public List<ChartDataPoint> InventoryDonutData { get; set; } = new();
        public List<MakerLoadChartPoint> MakerLoadChartData { get; set; } = new();
        public List<DeliveryTimelineChartPoint> DeliveryTimelineChartData { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoading = true;
                Customers = await DashboardService.GetCustomersAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load customers: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task OnCustomerChanged()
        {
            try
            {
                SelectedArticle = null;
                DashboardData = null;
                FilterOrderNo = null;
                SelectedStageFilter = null;
                ClearChartData();

                Countries = await DashboardService.GetCountriesByCustomerAsync(SelectedCustCode);
                SelectedCountry = Countries.FirstOrDefault() ?? "<Comp>";
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load countries: {ex.Message}");
            }
        }

        public Task OnCountryChanged()
        {
            SelectedArticle = null;
            DashboardData = null;
            FilterOrderNo = null;
            SelectedStageFilter = null;
            ClearChartData();
            return Task.CompletedTask;
        }

        public async Task<IEnumerable<LookupItemString>> SearchArticles(string searchText)
        {
            try
            {
                var list = await DashboardService.GetArticlesAsync(SelectedCustCode, SelectedCountry);
                if (string.IsNullOrWhiteSpace(searchText))
                    return list.Take(50);

                return list.Where(x => x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                                       x.Id.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                           .Take(50);
            }
            catch
            {
                return Enumerable.Empty<LookupItemString>();
            }
        }

        public async Task OnArticleChanged()
        {
            FilterOrderNo = null;
            SelectedStageFilter = null;
            await RefreshDashboardAsync();
        }

        public async Task RefreshDashboardAsync()
        {
            if (SelectedArticle == null || string.IsNullOrWhiteSpace(SelectedArticle.Id))
            {
                DashboardData = null;
                ClearChartData();
                return;
            }

            try
            {
                IsLoading = true;
                DashboardData = await DashboardService.GetDashboardDataAsync(SelectedArticle.Id, FilterOrderNo);
                BuildChartProjections();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load dashboard data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ClearChartData()
        {
            DemandCoveragePercent = 0;
            StageWipChartData.Clear();
            InventoryDonutData.Clear();
            MakerLoadChartData.Clear();
            DeliveryTimelineChartData.Clear();
        }

        private void BuildChartProjections()
        {
            ClearChartData();
            if (DashboardData == null) return;

            var sum = DashboardData.Summary;

            // 1. Demand Coverage Ratio (Total Supply / Total Demand * 100)
            if (sum.EffectiveDemand > 0)
            {
                DemandCoveragePercent = Math.Round(((double)sum.TotalSuppliesInPipeline / sum.EffectiveDemand) * 100.0, 1);
            }
            else
            {
                DemandCoveragePercent = sum.TotalSuppliesInPipeline > 0 ? 100.0 : 0;
            }

            // 2. Shop Floor WIP Bottleneck Distribution (Clean percentages and names)
            double totalWip = sum.TotalRunningLots > 0 ? sum.TotalRunningLots : 1;
            var wipByStage = DashboardData.RunningLots
                .GroupBy(x => string.IsNullOrWhiteSpace(x.Description) ? "General Processing" : x.Description.Trim())
                .Select(g => new StageWipChartPoint
                {
                    StageName = g.Key,
                    Quantity = g.Sum(x => x.Qty),
                    LotCount = g.Count(),
                    ProcessID = g.FirstOrDefault()?.ProcessID ?? 0,
                    PercentageOfTotal = Math.Round(((double)g.Sum(x => x.Qty) / totalWip) * 100.0, 1)
                })
                .OrderByDescending(x => x.Quantity)
                .ToList();
            StageWipChartData = wipByStage;

            // 3. Multi-Tier Inventory Donut Composition
            var invList = new List<ChartDataPoint>();
            if (sum.TotalRFStock > 0)
                invList.Add(new ChartDataPoint { Category = "Ready Finish", Value = sum.TotalRFStock, Color = "#107C41" });
            if (DashboardData.ReadyFinishLookalikes.Any())
                invList.Add(new ChartDataPoint { Category = "RF Lookalikes", Value = DashboardData.ReadyFinishLookalikes.Sum(x => x.NetQty), Color = "#2E7D32" });
            if (sum.TotalSFStock > 0)
                invList.Add(new ChartDataPoint { Category = "Semi-Finish", Value = sum.TotalSFStock, Color = "#00BFA5" });
            if (DashboardData.SemiFinishLookalikes.Any())
                invList.Add(new ChartDataPoint { Category = "SF Lookalikes", Value = DashboardData.SemiFinishLookalikes.Sum(x => x.Qty), Color = "#00897B" });
            if (sum.TotalForging > 0)
                invList.Add(new ChartDataPoint { Category = "Forgings / RM", Value = sum.TotalForging, Color = "#E65100" });
            InventoryDonutData = invList;

            // 4. Maker Subcontractor Load Distribution
            double totalPoQty = sum.TotalOpenPO > 0 ? sum.TotalOpenPO : 1;
            var makerGroups = DashboardData.OpenPOs
                .GroupBy(x => string.IsNullOrWhiteSpace(x.MakerName) ? "Unassigned Maker" : x.MakerName.Trim())
                .Select(g => new MakerLoadChartPoint
                {
                    MakerName = g.Key,
                    Quantity = g.Sum(x => x.Qty),
                    PoCount = g.Count(),
                    HasOverdue = g.Any(x => x.IsOverdue),
                    OverdueCount = g.Count(x => x.IsOverdue),
                    PercentageOfTotal = Math.Round(((double)g.Sum(x => x.Qty) / totalPoQty) * 100.0, 1)
                })
                .OrderByDescending(x => x.Quantity)
                .ToList();
            MakerLoadChartData = makerGroups;

            // 5. Delivery Timeline Schedule Urgency
            var today = DateTime.Today;
            double effectiveDemand = sum.TotalCustOrders > 0 ? sum.TotalCustOrders : 1;

            var overdueOrders = DashboardData.CustomerOrders.Where(x => x.DeliveryDT.HasValue && x.DeliveryDT.Value.Date < today).ToList();
            var next30Days = DashboardData.CustomerOrders.Where(x => x.DeliveryDT.HasValue && x.DeliveryDT.Value.Date >= today && x.DeliveryDT.Value.Date <= today.AddDays(30)).ToList();
            var next60Days = DashboardData.CustomerOrders.Where(x => x.DeliveryDT.HasValue && x.DeliveryDT.Value.Date > today.AddDays(30) && x.DeliveryDT.Value.Date <= today.AddDays(60)).ToList();
            var beyond60Days = DashboardData.CustomerOrders.Where(x => !x.DeliveryDT.HasValue || x.DeliveryDT.Value.Date > today.AddDays(60)).ToList();

            var timelineList = new List<DeliveryTimelineChartPoint>();
            if (overdueOrders.Any())
            {
                long qty = overdueOrders.Sum(x => x.UnshippedQty);
                timelineList.Add(new DeliveryTimelineChartPoint
                {
                    TimelineBucket = "Overdue / Critical",
                    BadgeClass = "bg-danger text-white",
                    IconClass = "fas fa-triangle-exclamation text-danger",
                    Quantity = qty,
                    OrderCount = overdueOrders.Count,
                    PercentageOfDemand = Math.Round(((double)qty / effectiveDemand) * 100.0, 1)
                });
            }
            if (next30Days.Any())
            {
                long qty = next30Days.Sum(x => x.UnshippedQty);
                timelineList.Add(new DeliveryTimelineChartPoint
                {
                    TimelineBucket = "Next 30 Days",
                    BadgeClass = "bg-warning text-dark",
                    IconClass = "fas fa-bolt text-warning",
                    Quantity = qty,
                    OrderCount = next30Days.Count,
                    PercentageOfDemand = Math.Round(((double)qty / effectiveDemand) * 100.0, 1)
                });
            }
            if (next60Days.Any())
            {
                long qty = next60Days.Sum(x => x.UnshippedQty);
                timelineList.Add(new DeliveryTimelineChartPoint
                {
                    TimelineBucket = "31 to 60 Days",
                    BadgeClass = "bg-info text-dark",
                    IconClass = "fas fa-calendar-days text-info",
                    Quantity = qty,
                    OrderCount = next60Days.Count,
                    PercentageOfDemand = Math.Round(((double)qty / effectiveDemand) * 100.0, 1)
                });
            }
            if (beyond60Days.Any())
            {
                long qty = beyond60Days.Sum(x => x.UnshippedQty);
                timelineList.Add(new DeliveryTimelineChartPoint
                {
                    TimelineBucket = "60+ Days / Open",
                    BadgeClass = "bg-secondary text-white",
                    IconClass = "fas fa-clock text-secondary",
                    Quantity = qty,
                    OrderCount = beyond60Days.Count,
                    PercentageOfDemand = Math.Round(((double)qty / effectiveDemand) * 100.0, 1)
                });
            }
            DeliveryTimelineChartData = timelineList;
        }

        // Cross-Filtering Interactions
        public async Task OnOrderRowClicked(string orderNo)
        {
            if (FilterOrderNo == orderNo)
            {
                FilterOrderNo = null;
            }
            else
            {
                FilterOrderNo = orderNo;
            }
            await RefreshDashboardAsync();
        }

        public async Task ClearOrderFilter()
        {
            FilterOrderNo = null;
            await RefreshDashboardAsync();
        }

        public void FilterByStage(string stageName)
        {
            if (SelectedStageFilter == stageName)
                SelectedStageFilter = null;
            else
            {
                SelectedStageFilter = stageName;
                ActiveMatrixTab = 1; // Switch matrix tab to WIP Lots
            }
        }

        public void ClearStageFilter()
        {
            SelectedStageFilter = null;
        }

        public IEnumerable<RunningLotItem> GetFilteredRunningLots()
        {
            if (DashboardData == null) return Enumerable.Empty<RunningLotItem>();
            if (string.IsNullOrWhiteSpace(SelectedStageFilter))
                return DashboardData.RunningLots;

            return DashboardData.RunningLots.Where(x =>
                string.Equals(x.Description, SelectedStageFilter, StringComparison.OrdinalIgnoreCase));
        }

        public async Task OpenCrossDemandModal(long rmid, string materialCode)
        {
            if (rmid == 0 || SelectedArticle == null) return;

            try
            {
                ActiveForgingCode = materialCode;
                CrossDemandItems = await DashboardService.GetForgingCrossDemandAsync(rmid, SelectedArticle.Id);
                ShowCrossDemandModal = true;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to fetch competing demand: {ex.Message}");
            }
        }

        public async Task ToggleForgingStatus(ForgingRawMaterialItem? forg)
        {
            if (forg == null || SelectedArticle == null) return;

            try
            {
                bool newInactive = !forg.IsInactive;
                bool ok = await DashboardService.ToggleForgingFunctionalStatusAsync(SelectedArticle.Id, forg.RMID, newInactive);
                if (ok)
                {
                    forg.IsInactive = newInactive;
                    NotificationService.Notify(NotificationSeverity.Success, "Updated", $"Forging status updated to {(newInactive ? "Inactive" : "Active")}.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to update forging status: {ex.Message}");
            }
        }

        public async Task CloseLotPrompt(string? lotNo)
        {
            if (string.IsNullOrWhiteSpace(lotNo)) return;

            bool? confirmed = await DialogService.Confirm($"Are you sure you want to close Lot #{lotNo}?", "Close Production Lot",
                new ConfirmOptions { OkButtonText = "Yes, Close Lot", CancelButtonText = "Cancel" });

            if (confirmed == true)
            {
                try
                {
                    bool ok = await DashboardService.CloseLotAsync(lotNo, "User", Environment.MachineName);
                    if (ok)
                    {
                        NotificationService.Notify(NotificationSeverity.Success, "Closed", $"Lot #{lotNo} has been closed.");
                        await RefreshDashboardAsync();
                    }
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to close lot: {ex.Message}");
                }
            }
        }

        public async Task CloseMakerPOPrompt(long entryId)
        {
            if (entryId == 0) return;

            bool? confirmed = await DialogService.Confirm("Are you sure you want to close this Maker Purchase Order?", "Close Maker PO",
                new ConfirmOptions { OkButtonText = "Yes, Close PO", CancelButtonText = "Cancel" });

            if (confirmed == true)
            {
                try
                {
                    bool ok = await DashboardService.ClosePOAsync(entryId);
                    if (ok)
                    {
                        NotificationService.Notify(NotificationSeverity.Success, "Closed", "Maker PO has been closed.");
                        await RefreshDashboardAsync();
                    }
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to close PO: {ex.Message}");
                }
            }
        }

        // ── REPORT & ACTION HANDLERS ─────────────────────────────────────────

        public async Task PrintSpecs()
        {
            if (SelectedArticle == null) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "ItemSpecification.rpt",
                    SelectionFormula = $"{{Items.ItemID}}='{SelectedArticle.Id}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintDetailSpecs()
        {
            if (SelectedArticle == null) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "ItemSpecification_Detail.rpt",
                    SelectionFormula = $"{{Items.ItemID}}='{SelectedArticle.Id}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintOrderBalance(string? orderNo)
        {
            if (string.IsNullOrWhiteSpace(orderNo)) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptProductionOrder_Balance.rpt",
                    SelectionFormula = $"{{CustOrders.OrderNo}}='{orderNo}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintOrderPlanningReportR1(string? orderNo, bool r1)
        {
            if (string.IsNullOrWhiteSpace(orderNo)) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "OrderPlanningReport.rpt",
                    SelectionFormula = $"{{CustOrders.OrderNo}}='{orderNo}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintOrderStatus(string? orderNo)
        {
            if (string.IsNullOrWhiteSpace(orderNo)) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptProductionOrder_Status.rpt",
                    Parameters = new Dictionary<string, object> { { "@OrderNo", orderNo } }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintPTC(string? lotNo, bool mini)
        {
            if (string.IsNullOrWhiteSpace(lotNo)) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = mini ? "PTC_Mini.rpt" : "PTCQEL.rpt",
                    SelectionFormula = $"{{VFOrderItemPTC.LotNo}}='{lotNo}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintRunningLotsReport()
        {
            if (SelectedArticle == null) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "PPDashBoard_RunningLots.rpt",
                    SelectionFormula = $"{{VRunningLots_Simple.ItemCode}}='{SelectedArticle.Id}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintFinishStoreLots()
        {
            if (SelectedArticle == null) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "PPDashBoard_FinishStoreLots.rpt",
                    SelectionFormula = $"{{VItemsWithShelfWiseStock.ItemID}}='{SelectedArticle.Id}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintMakerPO(string? masterPONo)
        {
            if (string.IsNullOrWhiteSpace(masterPONo)) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "IssList.rpt",
                    SelectionFormula = $"{{VendIssued.MasterPONo}}='{masterPONo}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintMasterPOStatus(string? masterPONo)
        {
            if (string.IsNullOrWhiteSpace(masterPONo)) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "MasterPOStatus.rpt",
                    SelectionFormula = $"{{VendIssued.MasterPONo}}='{masterPONo}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintRMLinkedReport()
        {
            if (SelectedArticle == null) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "PPDashBoard_RMLinkedLots.rpt",
                    SelectionFormula = $"{{ItemsRMComp.ItemID}}='{SelectedArticle.Id}' AND {{VMaterialLocationWiseBalances.ShelfQty}}>0"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public async Task PrintComplaint(long entryId)
        {
            if (entryId == 0) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "CustomerComplaints.rpt",
                    SelectionFormula = $"{{CustomerComplaints.EntryID}}={entryId}"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }

        public void ShowItemPicture()
        {
            if (SelectedArticle == null) return;
            NotificationService.Notify(NotificationSeverity.Info, "Item Preview", $"Loading picture for article {SelectedArticle.Id}...");
        }

        public void ShowTechnicalDrawing()
        {
            if (SelectedArticle == null) return;
            NotificationService.Notify(NotificationSeverity.Info, "Technical Drawing", $"Opening technical drawing for {SelectedArticle.Id}...");
        }

        public async Task OpenDurationModal()
        {
            if (SelectedArticle == null) return;
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "ItemWiseAvgProcessDuration.rpt",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@ItemID", SelectedArticle.Id },
                        { "@DTFrom", DateTime.Today.AddYears(-1) },
                        { "@DTTo", DateTime.Today }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
            }
        }
    }
}
