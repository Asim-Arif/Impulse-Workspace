using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels;
using Impulse.Services;
using Impulse.Services.Payroll;
using Impulse.Services.Production;
using Impulse.Shared.Components;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Production.Statistics
{
    public partial class ProductionStatisticsHub : ComponentBase
    {
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IEmployeeService EmployeeService { get; set; } = null!;
        [Inject] private IProcessService ProcessService { get; set; } = null!;

        private DateRangeWithTripleSelectModal productivityModal = null!;
        private DateRangeWithSubHeadsSelectModal deptEmpDateRangeModal = null!;
        private DateSingleWithSelectModal singleDeptModal = null!;
        private DateRangeModal rangeModal = null!;

        private List<GenericDropDownModel> departmentOptions = new();
        private List<GenericDropDownModel> allEmployeeOptions = new();
        private List<GenericDropDownModel> processOptions = new();

        private class StatCard
        {
            public string Name { get; set; } = string.Empty;
            public string Icon { get; set; } = string.Empty;
        }

        private List<StatCard> statistics = new()
        {
            new StatCard { Name = "Open Master POs", Icon = "fas fa-file-invoice" },
            new StatCard { Name = "Open Lots", Icon = "fas fa-cubes" },
            new StatCard { Name = "Emp. Productivity", Icon = "fas fa-user-check" },
            new StatCard { Name = "Emp. Productivity Monthly", Icon = "fas fa-calendar-alt" },
            new StatCard { Name = "Department Productivity", Icon = "fas fa-users-cog" },
            new StatCard { Name = "Productivity Chart", Icon = "fas fa-chart-area" },
            new StatCard { Name = "Running Lots Hubwise", Icon = "fas fa-dolly" },
            new StatCard { Name = "Running Lots", Icon = "fas fa-tasks" },
            new StatCard { Name = "Running Lots Valuewise", Icon = "fas fa-file-invoice-dollar" },
            new StatCard { Name = "Running Lots (Stock Order)", Icon = "fas fa-boxes-packing" },
            new StatCard { Name = "Lots Packed", Icon = "fas fa-box-open" },
            new StatCard { Name = "Production Item List", Icon = "fas fa-layer-group" }
        };

        protected override async Task OnInitializedAsync()
        {
            try
            {
                // 1. Departments (< All Departments > = "0")
                departmentOptions.Add(new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "< All Departments >" });
                var depts = await EmployeeService.GetDepartmentsAsync(false);
                if (depts != null)
                {
                    foreach (var d in depts)
                    {
                        departmentOptions.Add(new GenericDropDownModel
                        {
                            DropDownValue_ID = d.DeptID,
                            DropDownValue_Description = $"{d.Name} {{{d.DeptID}}}"
                        });
                    }
                }

                // 2. All Employees (<All Employees> = "0")
                allEmployeeOptions.Add(new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "<All Employees>" });
                var emps = await EmployeeService.GetEmployeesAsync(string.Empty, false, false);
                if (emps != null)
                {
                    foreach (var e in emps)
                    {
                        allEmployeeOptions.Add(new GenericDropDownModel
                        {
                            DropDownValue_ID = e.EmpID,
                            DropDownValue_Description = $"{{{e.EmpID}}} {e.Name}"
                        });
                    }
                }

                // 3. Processes (<All Processes> = "0")
                processOptions.Add(new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "<All Processes>" });
                var processes = await ProcessService.GetProcessesAsync(false);
                if (processes != null)
                {
                    foreach (var p in processes)
                    {
                        processOptions.Add(new GenericDropDownModel
                        {
                            DropDownValue_ID = p.ProcessID.ToString(),
                            DropDownValue_Description = p.Description
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Lookup Loading Notice",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        private async Task<List<GenericDropDownModel>> OnDepartmentChangedAsync(string deptId)
        {
            var result = new List<GenericDropDownModel>
            {
                new GenericDropDownModel { DropDownValue_ID = "0", DropDownValue_Description = "<All Employees>" }
            };

            if (string.IsNullOrEmpty(deptId) || deptId == "0")
            {
                return allEmployeeOptions;
            }

            try
            {
                var emps = await EmployeeService.GetEmployeesAsync(deptId, false, false);
                if (emps != null)
                {
                    foreach (var e in emps)
                    {
                        result.Add(new GenericDropDownModel
                        {
                            DropDownValue_ID = e.EmpID,
                            DropDownValue_Description = $"{{{e.EmpID}}} {e.Name}"
                        });
                    }
                }
            }
            catch
            {
                return allEmployeeOptions;
            }

            return result;
        }

        private async Task OnCardClick(string name)
        {
            try
            {
                if (name == "Open Master POs")
                {
                    await PrintOpenMakerOrdersAsync();
                }
                else if (name == "Open Lots")
                {
                    await PrintOpenLotsAsync();
                }
                else if (name == "Emp. Productivity")
                {
                    await PrintEmpProductivityReportAsync();
                }
                else if (name == "Emp. Productivity Monthly")
                {
                    await PrintEmpProductivityReportMonthlyAsync();
                }
                else if (name == "Department Productivity")
                {
                    await PrintDepartmentProductivityReportAsync();
                }
                else if (name == "Productivity Chart")
                {
                    await PrintProductivityChartReportAsync();
                }
                else if (name == "Running Lots Hubwise")
                {
                    await PrintRunningLotsHubwiseAsync();
                }
                else if (name == "Running Lots")
                {
                    await PrintRunningLotsAsync();
                }
                else if (name == "Running Lots Valuewise")
                {
                    await PrintRunningLotsValuewiseAsync();
                }
                else if (name == "Running Lots (Stock Order)")
                {
                    await PrintRunningLotsStockOrderAsync();
                }
                else if (name == "Lots Packed")
                {
                    await PrintLotsPackedAsync();
                }
                else if (name == "Production Item List")
                {
                    NavigationManager.NavigateTo("/production/item-list");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Execution Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task PrintOpenMakerOrdersAsync()
        {
            var req = new ReportRequest
            {
                ReportName = "Maker_Open_POs.rpt",
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>()
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintOpenLotsAsync()
        {
            var req = new ReportRequest
            {
                ReportName = "Open_Lots.rpt",
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>()
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintEmpProductivityReportAsync()
        {
            var res = await productivityModal.Show(
                departmentOptions,
                allEmployeeOptions,
                processOptions,
                "Employee Productivity Report"
            );

            if (!res.Success) return;

            var dtFrom = res.DateFrom;
            var dtTo = res.DateTo;
            var strDeptID = res.SelectedValue1 ?? "0";
            var strEmpID = res.SelectedValue2 ?? "0";
            var strProcesses = res.SelectedValue3 ?? "0";

            string selection = $"{{VEmpProductivityReport.DTOnly}}=#{dtFrom:yyyy-MM-dd}# TO #{dtTo:yyyy-MM-dd}#";

            if (strEmpID != "0")
            {
                selection += $" AND {{VEmpProductivityReport.EmpID}}='{strEmpID}'";
            }

            if (strDeptID != "0")
            {
                selection += $" AND {{VEmpProductivityReport.DeptID}}='{strDeptID}'";
            }

            if (strProcesses != "0")
            {
                selection += $" AND {{VEmpProductivityReport.ProcessID}} IN [{strProcesses.Replace("'", "")}]";
            }

            var req = new ReportRequest
            {
                ReportName = "EmployeeProductivityReportMerged.rpt",
                SelectionFormula = selection,
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>
                {
                    { "SubHeading", $"'{dtFrom:dd-MMM-yyyy} to {dtTo:dd-MMM-yyyy}'" },
                    { "ReportType", "'Regular'" },
                    { "DTFrom", $"#{dtFrom:yyyy-MM-dd}#" },
                    { "DTTo", $"#{dtTo:yyyy-MM-dd}#" }
                }
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintEmpProductivityReportMonthlyAsync()
        {
            var res = await deptEmpDateRangeModal.Show(
                departmentOptions,
                allEmployeeOptions,
                "Employee Monthly Productivity Report"
            );

            if (!res.Success) return;

            var dtFrom = res.DateFrom;
            var dtTo = res.DateTo;
            var strDeptID = res.SelectedValue ?? "0";
            var strEmpID = res.SelectedValue_sub ?? "0";

            string selection = $"{{VEmpProductivityReport_Production.DT}}=#{dtFrom:yyyy-MM-dd}# TO #{dtTo:yyyy-MM-dd}#";

            if (strEmpID != "0")
            {
                selection += $" AND {{VEmpProductivityReport_Production.EmpID}}='{strEmpID}'";
            }
            else if (strDeptID != "0")
            {
                selection += $" AND {{VEmpProductivityReport_Production.DeptiD}}='{strDeptID}'";
            }

            var req = new ReportRequest
            {
                ReportName = "EmployeeMonthlyProductivityReport.rpt",
                SelectionFormula = selection,
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>
                {
                    { "Range", $"'{dtFrom:dd-MMM-yyyy} to {dtTo:dd-MMM-yyyy}'" }
                }
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintDepartmentProductivityReportAsync()
        {
            var res = await singleDeptModal.Show(
                departmentOptions,
                "Department Productivity Report"
            );

            if (!res.Success) return;

            var dtFrom = res.Date;
            var strDeptID = res.SelectedValue ?? "0";

            string selection = $"{{VEmpProductivityReport.DTOnly}}=#{dtFrom:yyyy-MM-dd}# AND {{VEmpProductivityReport.OverTime}}=FALSE";

            if (strDeptID != "0" && !string.IsNullOrEmpty(strDeptID) && strDeptID != "All")
            {
                selection += $" AND {{VEmpProductivityReport.DeptID}} IN ['{strDeptID}']";
            }

            var req = new ReportRequest
            {
                ReportName = "EmployeeProductivityReportDatewise.rpt",
                SelectionFormula = selection,
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>
                {
                    { "SubHeading", $"'{dtFrom:dd-MMM-yyyy}'" },
                    { "ReportType", "'Regular'" }
                }
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintProductivityChartReportAsync()
        {
            var res = await productivityModal.Show(
                departmentOptions,
                allEmployeeOptions,
                processOptions,
                "Productivity Chart Report"
            );

            if (!res.Success) return;

            var dtFrom = res.DateFrom;
            var dtTo = res.DateTo;

            string selection = $"{{VEmpProductivityReport_Production.DTOnly}}=#{dtFrom:yyyy-MM-dd}# TO #{dtTo:yyyy-MM-dd}# AND {{VEmpProductivityReport_Production.OverTime}}=FALSE";

            var req = new ReportRequest
            {
                ReportName = "EmployeeProductivityReportLineChart.rpt",
                SelectionFormula = selection,
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>
                {
                    { "DateRange", $"'{dtFrom:dd-MMM-yyyy} to {dtTo:dd-MMM-yyyy}'" },
                    { "ReportType", "'Regular'" },
                    { "DTFrom", $"#{dtFrom:yyyy-MM-dd}#" },
                    { "DTTo", $"#{dtTo:yyyy-MM-dd}#" }
                }
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintRunningLotsHubwiseAsync()
        {
            var res = await rangeModal.Show("Running Lots Summary Hubwise");
            if (!res.Success) return;

            var req = new ReportRequest
            {
                ReportName = "Running_Lots_Summary_Hubwise.rpt",
                Parameters = new Dictionary<string, object>
                {
                    { "@DTFrom", res.DateFrom.ToString("yyyy-MM-dd") },
                    { "@DTTo", res.DateTo.ToString("yyyy-MM-dd") },
                    { "@OrderNo", "0" }
                },
                FormulaValues = new Dictionary<string, object>()
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintRunningLotsAsync()
        {
            var req = new ReportRequest
            {
                ReportName = "RunningLots.rpt",
                SelectionFormula = string.Empty,
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>()
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintRunningLotsValuewiseAsync()
        {
            var req = new ReportRequest
            {
                ReportName = "RunningLots_Value.rpt",
                SelectionFormula = string.Empty,
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>()
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintRunningLotsStockOrderAsync()
        {
            var req = new ReportRequest
            {
                ReportName = "RunningLotsStock.rpt",
                SelectionFormula = "{VRunningLots.OrderNo} IN ['Stock-Order', 'Stock-OrderIss']",
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>()
            };

            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task PrintLotsPackedAsync()
        {
            var req = new ReportRequest
            {
                ReportName = "LotsPacked.rpt",
                SelectionFormula = "{VRunningLots.ProcessID} IN [163]",
                Parameters = new Dictionary<string, object>(),
                FormulaValues = new Dictionary<string, object>()
            };

            await ReportNavigation.PrintReportAsync(req);
        }
    }
}
