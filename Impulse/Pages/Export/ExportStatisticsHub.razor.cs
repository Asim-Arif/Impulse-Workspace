using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Impulse.Services;
using Impulse.Shared.Components;
using DataAccessLibrary.Models.ViewModels;

namespace Impulse.Pages.Export
{
    public partial class ExportStatisticsHub : ComponentBase
    {
        [Inject] private INotificationService Notification { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private Impulse.Services.Export.ICustomerOrderService CustomerOrderService { get; set; } = null!;
        [Inject] private Impulse.Services.Export.IExportPerformanceReportService ExportPerformanceReportService { get; set; } = null!;

        private DateSingleModal singleModal = null!;
        private DateRangeModal rangeModal = null!;
        private DateRangeWithSingleSelectModal customerDateRangeModal = null!;
        private DateRangeWithSubHeadsSelectModal itemGroupsDateRangeModal = null!;

        private List<GenericDropDownModel> customerOptions = new();
        private List<GenericDropDownModel> itemGroupOptions = new();
        private List<GenericDropDownModel> additionalGroupOptions = new();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var fetchedCustomers = await CustomerOrderService.GetCustomersAsync();
                customerOptions.Add(new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "All Customers" });
                if (fetchedCustomers != null)
                {
                    foreach (var cust in fetchedCustomers)
                    {
                        customerOptions.Add(new GenericDropDownModel
                        {
                            DropDownValue_ID = cust.CustCode,
                            DropDownValue_Description = $"{cust.CustCode} - {cust.Name}"
                        });
                    }
                }

                // Load item groups
                var fetchedItemGroups = await ExportPerformanceReportService.GetItemGroupsAsync();
                itemGroupOptions.Add(new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "<All Groups>" });
                if (fetchedItemGroups != null)
                {
                    foreach (var grp in fetchedItemGroups)
                    {
                        itemGroupOptions.Add(new GenericDropDownModel
                        {
                            DropDownValue_ID = grp.ID.ToString(),
                            DropDownValue_Description = grp.Description
                        });
                    }
                }

                // Load main (additional) groups
                var fetchedMainGroups = await ExportPerformanceReportService.GetMainGroupsAsync();
                additionalGroupOptions.Add(new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "<All Groups>" });
                if (fetchedMainGroups != null)
                {
                    foreach (var grp in fetchedMainGroups)
                    {
                        additionalGroupOptions.Add(new GenericDropDownModel
                        {
                            DropDownValue_ID = grp.MainGroupID.ToString(),
                            DropDownValue_Description = grp.MainGroupName
                        });
                    }
                }
            }
            catch (System.Exception ex)
            {
                Notification.ShowError("Initialization Error", "Failed to load lookup data: " + ex.Message);
            }
        }



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
            new StatisticItem { Name = "Main Group/Sub-Group Item Level D&S", Icon = "fas fa-indent" },
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
            else if (name == "Main Group/Sub-Group Item Level D&S")
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
            else if (name == "Customer Purchase History")
            {
                if (customerOptions == null || customerOptions.Count <= 1)
                {
                    try
                    {
                        var fetchedCustomers = await CustomerOrderService.GetCustomersAsync();
                        customerOptions = new List<GenericDropDownModel>
                        {
                            new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "All Customers" }
                        };
                        if (fetchedCustomers != null)
                        {
                            foreach (var cust in fetchedCustomers)
                            {
                                customerOptions.Add(new GenericDropDownModel
                                {
                                    DropDownValue_ID = cust.CustCode,
                                    DropDownValue_Description = $"{cust.CustCode} - {cust.Name}"
                                });
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Notification.ShowError("Error", "Failed to load customers: " + ex.Message);
                        return;
                    }
                }

                var res = await customerDateRangeModal.Show(customerOptions, "Customer Purchase History");
                if (res.Success)
                {
                    var selectionFormula = $"{{FCustomerOrders.DT}} >= #{res.DateFrom:yyyy-MM-dd}# AND {{FCustomerOrders.DT}} <= #{res.DateTo:yyyy-MM-dd}#";
                    if (res.SelectedValue != "0" && !string.IsNullOrEmpty(res.SelectedValue))
                    {
                        selectionFormula += $" AND {{FCustomerOrders.CustCode}}='{res.SelectedValue}'";
                    }

                    var req = new ReportRequest
                    {
                        ReportName = "CustomerPurchaseHistory_Orders.rpt",
                        SelectionFormula = selectionFormula,
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{res.DateFrom:dd-MMM-yyyy} to {res.DateTo:dd-MMM-yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Customer Purchase History report dispatched.");
                }
            }
            else if (name == "Itemwise Yearly Demand")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "YearlyItemwiseDemandReport.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DT", res.Date },
                            { "@YearlyReport", false },
                            { "@X2Only", false }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "Year", $"'{res.Date:yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Itemwise yearly demand report dispatched.");
                }
            }
            else if (name == "Itemwise Yearly Demand (X2)")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "YearlyItemwiseDemandReport.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DT", res.Date },
                            { "@YearlyReport", false },
                            { "@X2Only", true }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "Year", $"'{res.Date:yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Itemwise yearly demand report (X2) dispatched.");
                }
            }
            else if (name == "Itemwise Yearwise Demand")
            {
                var res = await singleModal.ShowYearOnly();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "YearlyItemwiseDemandReport_Yearwise.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DT", res.Date },
                            { "@YearlyReport", true },
                            { "@X2Only", false }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "Year", $"'{res.Date:yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Itemwise yearwise demand report dispatched.");
                }
            }
            else if (name == "Yearly Customerwise Demands")
            {
                var req = new ReportRequest
                {
                    ReportName = "CustomerWiseDemands_Yearwise.rpt",
                    SelectionFormula = "{VFOrderList.OrderType}='Customer Order' AND {VFOrderList.Cancelled}=0 AND (NOT ({VFOrderList.CustCode} IN['Stock','Wrangler']))"
                };
                await ReportNavigation.PrintReportAsync(req);
                Notification.ShowSuccess("Success", "Yearly customerwise demands report dispatched.");
            }
            else if (name == "Product Line Demand")
            {
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var selectionFormula = $"{{FCustomerOrders.DT}} >= #{res.DateFrom:yyyy-MM-dd}# AND {{FCustomerOrders.DT}} <= #{res.DateTo:yyyy-MM-dd}# " +
                                           $"AND (NOT ({{FCustomerOrders.CustCode}} IN['Stock','Wrangler','SAL'])) " +
                                           $"AND {{FCustomerOrders.OrderType}}='Customer Order' " +
                                           $"AND ({{FCustomerFinalOrders.Cancelled}}=0 OR ISNULL({{FCustomerFinalOrders.Cancelled}}))";

                    var req = new ReportRequest
                    {
                        ReportName = "ProductLineDemand.rpt",
                        SelectionFormula = selectionFormula,
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{res.DateFrom:dd-MMM-yyyy} to {res.DateTo:dd-MMM-yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Product Line Demand report dispatched.");
                }
            }
            else if (name == "Product Groupwise Demand")
            {
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var selectionFormula = $"{{FCustomerOrders.DT}} >= #{res.DateFrom:yyyy-MM-dd}# AND {{FCustomerOrders.DT}} <= #{res.DateTo:yyyy-MM-dd}# " +
                                           $"AND (NOT ({{FCustomerOrders.CustCode}} IN['Stock','Wrangler','SAL'])) " +
                                           $"AND {{FCustomerOrders.OrderType}}='Customer Order' " +
                                           $"AND ({{FCustomerFinalOrders.Cancelled}}=0 OR ISNULL({{FCustomerFinalOrders.Cancelled}}))";

                    var req = new ReportRequest
                    {
                        ReportName = "ProductGroupwiseDemand.rpt",
                        SelectionFormula = selectionFormula,
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{res.DateFrom:dd-MMM-yyyy} to {res.DateTo:dd-MMM-yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Product Groupwise Demand report dispatched.");
                }
            }
            else if (name == "Product Line Supply")
            {
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var selectionFormula = $"{{CustomInvoice.DT}} >= #{res.DateFrom:yyyy-MM-dd}# AND {{CustomInvoice.DT}} <= #{res.DateTo:yyyy-MM-dd}# " +
                                           $"AND {{CustomInvoice.GatePassNo}} <> ''";

                    var req = new ReportRequest
                    {
                        ReportName = "ProductLineSupply.rpt",
                        SelectionFormula = selectionFormula,
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{res.DateFrom:dd-MMM-yyyy} to {res.DateTo:dd-MMM-yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Product Line Supply report dispatched.");
                }
            }
            else if (name == "Product Groupwise Supply")
            {
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var selectionFormula = $"{{CustomInvoice.DT}} >= #{res.DateFrom:yyyy-MM-dd}# AND {{CustomInvoice.DT}} <= #{res.DateTo:yyyy-MM-dd}# " +
                                           $"AND {{CustomInvoice.GatePassNo}} <> ''";

                    var req = new ReportRequest
                    {
                        ReportName = "ProductGroupwiseSupply.rpt",
                        SelectionFormula = selectionFormula,
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{res.DateFrom:dd-MMM-yyyy} to {res.DateTo:dd-MMM-yyyy}'" }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Product Groupwise Supply report dispatched.");
                }
            }
            else if (name == "Total Demand Value Share")
            {
                // Call ItemWiseDemandSupply(True, 6)  →  bCustomerwise=True, iReportType=6
                var res = await rangeModal.Show();
                if (res.Success)
                {
                    var req = new ReportRequest
                    {
                        ReportName = "TotalOrderValueShare_Demand.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", res.DateFrom },
                            { "@DTTo",   res.DateTo }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Total Demand Value Share report dispatched.");
                }
            }
            else if (name == "Advance Plan")
            {
                if (itemGroupOptions == null || itemGroupOptions.Count <= 1 || additionalGroupOptions == null || additionalGroupOptions.Count <= 1)
                {
                    try
                    {
                        var fetchedItemGroups = await ExportPerformanceReportService.GetItemGroupsAsync();
                        itemGroupOptions = new List<GenericDropDownModel>
                        {
                            new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "<All Groups>" }
                        };
                        if (fetchedItemGroups != null)
                        {
                            foreach (var grp in fetchedItemGroups)
                            {
                                itemGroupOptions.Add(new GenericDropDownModel
                                {
                                    DropDownValue_ID = grp.ID.ToString(),
                                    DropDownValue_Description = grp.Description
                                });
                            }
                        }

                        var fetchedMainGroups = await ExportPerformanceReportService.GetMainGroupsAsync();
                        additionalGroupOptions = new List<GenericDropDownModel>
                        {
                            new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "<All Groups>" }
                        };
                        if (fetchedMainGroups != null)
                        {
                            foreach (var grp in fetchedMainGroups)
                            {
                                additionalGroupOptions.Add(new GenericDropDownModel
                                {
                                    DropDownValue_ID = grp.MainGroupID.ToString(),
                                    DropDownValue_Description = grp.MainGroupName
                                });
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Notification.ShowError("Error", "Failed to load lookup data: " + ex.Message);
                        return;
                    }
                }

                var res = await itemGroupsDateRangeModal.Show(itemGroupOptions, additionalGroupOptions, "Advance Plan");
                if (res.Success)
                {
                    var dtFrom = new DateTime(res.DateFrom.Year, 1, 1);
                    var dtTo = new DateTime(res.DateTo.Year, 12, 31);

                    var req = new ReportRequest
                    {
                        ReportName = "YearlyItemwiseDemandReport_EachYear.rpt",
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", dtFrom },
                            { "@DTTo", dtTo },
                            { "@ItemGroups", res.SelectedValue == "0" ? "" : (res.SelectedValue ?? "") },
                            { "@AdditionalGroups", res.SelectedValue_sub == "0" ? "" : (res.SelectedValue_sub ?? "") }
                        }
                    };
                    await ReportNavigation.PrintReportAsync(req);
                    Notification.ShowSuccess("Success", "Advance Plan report dispatched.");
                }
            }
            else
            {
                Notification.ShowInformation("Statistics", $"The '{name}' report has not been migrated yet.");
            }
        }
    }
}

