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
    public partial class ProductionPlanningDashboard : ComponentBase
    {
        [Inject] public IProductionPlanningDashboardService DashboardService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;
        [Inject] public DialogService DialogService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;

        public List<LookupItemString> Customers { get; set; } = new();
        public string SelectedCustCode { get; set; } = "0";

        public List<string> Countries { get; set; } = new() { "<Comp>" };
        public string SelectedCountry { get; set; } = "<Comp>";

        public LookupItemString? SelectedArticle { get; set; }
        public string? FilterOrderNo { get; set; }

        public ProductionPlanningDashboardData? DashboardData { get; set; }
        public bool IsLoading { get; set; } = false;

        public int InventoryTab { get; set; } = 0;
        public int DetailsTab { get; set; } = 0;

        public bool ShowCrossDemandModal { get; set; } = false;
        public string ActiveForgingCode { get; set; } = string.Empty;
        public List<ForgingCrossDemandItem> CrossDemandItems { get; set; } = new();

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
            await RefreshDashboardAsync();
        }

        public async Task RefreshDashboardAsync()
        {
            if (SelectedArticle == null || string.IsNullOrWhiteSpace(SelectedArticle.Id))
            {
                DashboardData = null;
                return;
            }

            try
            {
                IsLoading = true;
                DashboardData = await DashboardService.GetDashboardDataAsync(SelectedArticle.Id, FilterOrderNo);
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

        public async Task FilterByOrderNoFromContext(CustOrderDemandItem? order)
        {
            if (order == null) return;
            FilterOrderNo = order.OrderNo;
            await RefreshDashboardAsync();
        }

        public async Task ClearOrderFilter()
        {
            FilterOrderNo = null;
            await RefreshDashboardAsync();
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
