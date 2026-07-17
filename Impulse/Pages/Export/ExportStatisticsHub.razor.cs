using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Impulse.Services;
using Impulse.Shared.Components;

namespace Impulse.Pages.Export
{
    public partial class ExportStatisticsHub : ComponentBase
    {
        [Inject] private INotificationService Notification { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;

        private DateSingleModal singleModal = null!;
        private DateRangeModal rangeModal = null!;

        public class StatisticItem
        {
            public string Name { get; set; } = string.Empty;
            public string Icon { get; set; } = string.Empty;
        }

        private readonly List<StatisticItem> statistics = new()
        {
            new StatisticItem { Name = "Total Export", Icon = "fas fa-ship" },
            new StatisticItem { Name = "Total Demand", Icon = "fas fa-chart-line" },
            new StatisticItem { Name = "Articlewise Sales", Icon = "fas fa-chart-pie" },
            new StatisticItem { Name = "Articlewise Demand", Icon = "fas fa-chart-bar" },
            new StatisticItem { Name = "Itemwise Export Performance", Icon = "fas fa-tachometer-alt" },
            new StatisticItem { Name = "Customer Demand Report Current Year (Qty)", Icon = "fas fa-boxes" },
            new StatisticItem { Name = "Customer Demand Report Current Year (Amt)", Icon = "fas fa-dollar-sign" },
            new StatisticItem { Name = "Current Year monthwise demand", Icon = "fas fa-calendar-alt" },
            new StatisticItem { Name = "Yearly Demand/Supply", Icon = "fas fa-history" },
            new StatisticItem { Name = "Demand/Supply Comparison Graph Qty wise", Icon = "fas fa-chart-area" },
            new StatisticItem { Name = "Demand/Supply Comparison Graph Amt wise", Icon = "fas fa-chart-line" },
            new StatisticItem { Name = "Main Group Demand/Supply", Icon = "fas fa-folder" },
            new StatisticItem { Name = "Sub-Group Demand/Supply", Icon = "fas fa-folder-open" },
            new StatisticItem { Name = "Main Group/Sub-Group Item Level D&&S", Icon = "fas fa-indent" },
            new StatisticItem { Name = "Customerwise Demands", Icon = "fas fa-users" },
            new StatisticItem { Name = "Total Export Value Share", Icon = "fas fa-percentage" },
            new StatisticItem { Name = "Quarterly Demand", Icon = "fas fa-calendar-check" },
            new StatisticItem { Name = "Delivery Performance Report", Icon = "fas fa-truck" },
            new StatisticItem { Name = "Delivery Performance Detail", Icon = "fas fa-info-circle" },
            new StatisticItem { Name = "Customer Purchase History", Icon = "fas fa-history" },
            new StatisticItem { Name = "Itemwise Yearly Demand", Icon = "fas fa-list-ol" },
            new StatisticItem { Name = "Itemwise Yearly Demand (X2)", Icon = "fas fa-list-ul" },
            new StatisticItem { Name = "Itemwise Yearwise Demand", Icon = "fas fa-th-list" },
            new StatisticItem { Name = "Yearly Customerwise Demands", Icon = "fas fa-user-friends" },
            new StatisticItem { Name = "Product Line Demand", Icon = "fas fa-stream" },
            new StatisticItem { Name = "Product Groupwise Demand", Icon = "fas fa-layer-group" },
            new StatisticItem { Name = "Product Line Supply", Icon = "fas fa-truck-loading" },
            new StatisticItem { Name = "Product Groupwise Supply", Icon = "fas fa-network-wired" },
            new StatisticItem { Name = "Total Demand Value Share", Icon = "fas fa-chart-pie" },
            new StatisticItem { Name = "Advance Plan", Icon = "fas fa-calendar-plus" }
        };

        private async Task OnCardClick(string name)
        {
            if (name == "Total Export")
            {
                NavigationManager.NavigateTo("/export/statistics/total-export");
            }
            else if (name == "Total Demand")
            {
                NavigationManager.NavigateTo("/export/statistics/total-demand");
            }
            else if (name == "Articlewise Sales")
            {
                NavigationManager.NavigateTo("/export/statistics/articlewise-sales");
            }
            else if (name == "Articlewise Demand")
            {
                NavigationManager.NavigateTo("/export/statistics/articlewise-demand");
            }
            else if (name == "Itemwise Export Performance")
            {
                NavigationManager.NavigateTo("/export/statistics/itemwise-export-performance");
            }
            else if (name == "Customer Demand Report Current Year (Qty)")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "YearlyCustomerwiseDemandReport.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DT", res.Date }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "Year", $"'{res.Date:yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Yearly customerwise demand report dispatched.");
                }
            }
            else if (name == "Customer Demand Report Current Year (Amt)")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "YearlyCustomerwiseDemandReportAmtWise.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DT", res.Date }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "Year", $"'{res.Date:yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Yearly customerwise demand report (Amount-wise) dispatched.");
                }
            }
            else if (name == "Current Year monthwise demand")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "YearlyMainGroupwiseDemandReport.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DT", res.Date }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "Year", $"'{res.Date:yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Current year monthwise demand report dispatched.");
                }
            }
            else if (name == "Yearly Demand/Supply")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var dtFrom = new DateTime(res.Date.Year, 1, 1);
                    var dtTo   = new DateTime(res.Date.Year, 12, 31);
                    var req = new ReportRequest
                    {
                        ReportName = "DemandAndSupplyReport.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", dtFrom },
                            { "@DTTo",   dtTo }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{res.Date.Year}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Yearly Demand/Supply report dispatched.");
                }
            }
            else if (name == "Demand/Supply Comparison Graph Qty wise")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var dtFrom = new DateTime(res.Date.Year, 1, 1);
                    var dtTo   = new DateTime(res.Date.Year, 12, 31);
                    var req = new ReportRequest
                    {
                        ReportName = "DemandAndSupplyComparisonGraph.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", dtFrom },
                            { "@DTTo",   dtTo }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{res.Date.Year}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Demand/Supply Comparison Graph (Qty) dispatched.");
                }
            }
            else if (name == "Demand/Supply Comparison Graph Amt wise")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var dtFrom = new DateTime(res.Date.Year, 1, 1);
                    var dtTo   = new DateTime(res.Date.Year, 12, 31);
                    var req = new ReportRequest
                    {
                        ReportName = "DemandAndSupplyComparisonGraphAmtWise.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", dtFrom },
                            { "@DTTo",   dtTo }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{res.Date.Year}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Demand/Supply Comparison Graph (Amt) dispatched.");
                }
            }
            // --- ItemWiseDemandSupply family — all use DateRange (rangeModal) ---
            else if (name == "Main Group/Sub-Group Item Level D&&S")
            {
                // Call ItemWiseDemandSupply(0)  →  bCustomerwise=False, iReportType=0
                // ItemwiseDemandSupply.rpt  ShowItems=True, ShowSubGroup=True
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "ItemwiseDemandSupply.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", res.DateFrom },
                            { "@DTTo",   res.DateTo }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "ShowItems",    "True" },
                            { "ShowSubGroup", "True" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Main Group/Sub-Group Item Level D&S report dispatched.");
                }
            }
            else if (name == "Total Export Value Share")
            {
                // Call ItemWiseDemandSupply(True, 3)  →  bCustomerwise=True, iReportType=3
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "TotalOrderValueShare.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", res.DateFrom },
                            { "@DTTo",   res.DateTo }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Total Export Value Share report dispatched.");
                }
            }
            else if (name == "Delivery Performance Report")
            {
                // Call ItemWiseDemandSupply(False, 4)  →  ExportDeliveryPerformanceReport.rpt
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "ExportDeliveryPerformanceReport.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", res.DateFrom },
                            { "@DTTo",   res.DateTo }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Delivery Performance Report dispatched.");
                }
            }
            else if (name == "Delivery Performance Detail")
            {
                // Call ItemWiseDemandSupply(False, 5)  →  ExportDeliveryPerformanceDetailReport.rpt
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "ExportDeliveryPerformanceDetailReport.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", res.DateFrom },
                            { "@DTTo",   res.DateTo }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Delivery Performance Detail report dispatched.");
                }
            }
            else
            {
                Notification.ShowInformation("Statistics", $"The '{name}' report has not been migrated yet.");
            }
        }
    }
}
