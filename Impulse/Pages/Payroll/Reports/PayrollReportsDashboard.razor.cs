using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Impulse.Components.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Impulse.Shared.Components;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Payroll.Reports
{
    public partial class PayrollReportsDashboard : ComponentBase
    {
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private string SearchText { get; set; } = string.Empty;
        private List<PayrollReportCardItem> ReportItems { get; set; } = new();

        // Criteria Modal State
        private bool IsCriteriaModalOpen { get; set; } = false;
        private string CriteriaModalTitle { get; set; } = "Report Filter Criteria";
        private PayrollReportCardItem? ActiveReportItem { get; set; }

        // Shared Date Range Modal Reference
        private DateRangeModal DateRangeModalRef { get; set; } = default!;

        protected override void OnInitialized()
        {
            ReportItems = new List<PayrollReportCardItem>
            {
                new PayrollReportCardItem
                {
                    Id = "emp-list-with-salary",
                    Title = "Employee List (With Salary)",
                    Description = "Prints complete roster of active employees with full salary figures.",
                    Category = "Employee Master Reports",
                    Icon = "fas fa-users-rectangle",
                    BadgeText = "rptEmpList.rpt",
                    BadgeClass = "bg-primary",
                    ReportName = ReportNames.Payroll.EmpList,
                    SelectionFormula = "{VEmp.Active}=True  AND {VEmp.TempDept}=False"
                },
                new PayrollReportCardItem
                {
                    Id = "emp-list-without-salary",
                    Title = "Employee List (Without Salary)",
                    Description = "Prints complete roster of active employees excluding confidential salary information.",
                    Category = "Employee Master Reports",
                    Icon = "fas fa-address-book",
                    BadgeText = "rptEmpListWithoutSal.rpt",
                    BadgeClass = "bg-info text-dark",
                    ReportName = ReportNames.Payroll.EmpListWithoutSal,
                    SelectionFormula = "{VEmp.Active}=true  AND {VEmp.TempDept}=False"
                },
                new PayrollReportCardItem
                {
                    Id = "terminated-employees",
                    Title = "Terminated Employees",
                    Description = "Report of employees terminated / inactivated within specified date range.",
                    Category = "Employee Master Reports",
                    Icon = "fas fa-user-times",
                    BadgeText = "TerminatedEmployees.rpt",
                    BadgeClass = "bg-danger text-white",
                    ReportName = ReportNames.Payroll.TerminatedEmployees,
                    RequiresDateRangeModal = true
                },
                new PayrollReportCardItem
                {
                    Id = "new-hired",
                    Title = "New Hired",
                    Description = "Report of newly hired employees joined within specified date range.",
                    Category = "Employee Master Reports",
                    Icon = "fas fa-user-plus",
                    BadgeText = "NewHiredEmployees.rpt",
                    BadgeClass = "bg-success text-white",
                    ReportName = ReportNames.Payroll.NewHiredEmployees,
                    RequiresDateRangeModal = true
                },

                new PayrollReportCardItem
                {
                    Id = "salary-ledger",
                    Title = "Salary Ledger",
                    Description = "Detailed employee salary ledger statement by date range and department/employee.",
                    Category = "Salary Reports",
                    Icon = "fas fa-book-journal-whills",
                    BadgeText = "EmpSalaryLedger.rpt",
                    BadgeClass = "bg-warning text-dark",
                    ReportName = ReportNames.Payroll.EmpSalaryLedger,
                    RequiresFilterModal = true
                },
                new PayrollReportCardItem
                {
                    Id = "attendance-status",
                    Title = "Attendance Status",
                    Description = "Monthly employee attendance status summary report.",
                    Category = "Attendance & Absents",
                    Icon = "fas fa-calendar-check",
                    BadgeText = "MonthlyAttendanceStatus.rpt",
                    BadgeClass = "bg-primary",
                    ReportName = ReportNames.Payroll.MonthlyAttendanceStatus,
                    RequiresFilterModal = true
                },
                new PayrollReportCardItem
                {
                    Id = "employee-wise-leaves",
                    Title = "Employee wise Leaves",
                    Description = "Detailed employee leave log by date range and department/employee.",
                    Category = "Attendance & Absents",
                    Icon = "fas fa-plane-departure",
                    BadgeText = "LeavesList.rpt",
                    BadgeClass = "bg-success text-white",
                    ReportName = ReportNames.Payroll.LeavesList,
                    RequiresFilterModal = true
                },
                new PayrollReportCardItem
                {
                    Id = "missing-outtime",
                    Title = "Missing Outtime",
                    Description = "Employee attendance records with missing out time or single time entries.",
                    Category = "Attendance & Absents",
                    Icon = "fas fa-user-clock",
                    BadgeText = "MissingAttendance.rpt",
                    BadgeClass = "bg-danger text-white",
                    ReportName = ReportNames.Payroll.MissingAttendance,
                    RequiresFilterModal = true
                },
                new PayrollReportCardItem
                {
                    Id = "late-attendance",
                    Title = "Late Attendance",
                    Description = "Department-wise employee late attendance report and details.",
                    Category = "Attendance & Absents",
                    Icon = "fas fa-user-clock",
                    BadgeText = "LateAttReportDeptwise.rpt",
                    BadgeClass = "bg-warning text-dark",
                    ReportName = ReportNames.Payroll.LateAttReportDeptwise,
                    RequiresFilterModal = true
                },
                new PayrollReportCardItem
                {
                    Id = "employee-targets",
                    Title = "Employee Targets",
                    Description = "Employee targets and performance report by date range and employee.",
                    Category = "Performance & Targets",
                    Icon = "fas fa-bullseye",
                    BadgeText = "Employees_Targets_Performance.rpt",
                    BadgeClass = "bg-primary",
                    ReportName = ReportNames.Payroll.EmployeesTargetsPerformance,
                    RequiresFilterModal = true
                },


                new PayrollReportCardItem
                {
                    Id = "absent-sheet",
                    Title = "Absent Sheet",
                    Description = "Monthly employee attendance absent sheet report and details.",
                    Category = "Attendance & Absents",
                    Icon = "fas fa-user-slash",
                    NavigationUrl = "/payroll/absent-sheet?returnUrl=/payroll/reports"
                },
                new PayrollReportCardItem
                {
                    Id = "social-security",
                    Title = "Social Security",
                    Description = "Monthly social security contributions and forms.",
                    Category = "Statutory & Government",
                    Icon = "fas fa-shield-alt",
                    NavigationUrl = "/payroll/social-security?returnUrl=/payroll/reports"
                },
                new PayrollReportCardItem
                {
                    Id = "eobi-sheet",
                    Title = "Old Age",
                    Description = "EOBI monthly old age benefit contribution report.",
                    Category = "Statutory & Government",
                    Icon = "fas fa-building-user",
                    NavigationUrl = "/payroll/eobi?returnUrl=/payroll/reports"
                }
            };
        }

        private IEnumerable<PayrollReportCardItem> FilteredReports =>
            string.IsNullOrWhiteSpace(SearchText)
                ? ReportItems
                : ReportItems.Where(r => r.Title.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                                         r.Description.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                                         r.Category.Contains(SearchText, StringComparison.OrdinalIgnoreCase));

        private async Task ExecuteReportAsync(PayrollReportCardItem item)
        {
            if (!string.IsNullOrEmpty(item.NavigationUrl))
            {
                NavigationManager.NavigateTo(item.NavigationUrl);
                return;
            }

            if (item.RequiresDateRangeModal)
            {
                var res = await DateRangeModalRef.Show();
                if (!res.Success) return;

                try
                {
                    string strSelection = item.Id == "new-hired"
                        ? $"{{Employees.JoinDate}} in Date({res.DateFrom.Year}, {res.DateFrom.Month}, {res.DateFrom.Day}) to Date({res.DateTo.Year}, {res.DateTo.Month}, {res.DateTo.Day})"
                        : $"{{VEmpLastInActiveDate.InActiveDT}} in Date({res.DateFrom.Year}, {res.DateFrom.Month}, {res.DateFrom.Day}) to Date({res.DateTo.Year}, {res.DateTo.Month}, {res.DateTo.Day})";
                    var request = new ReportRequest
                    {
                        ReportName = item.ReportName,
                        SelectionFormula = strSelection,
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "MainHeading", $"'{res.DateFrom:dd-MMM-yyyy} To {res.DateTo:dd-MMM-yyyy}'" }
                        }
                    };

                    await ReportNavigationService.PrintReportAsync(request);

                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Info,
                        Summary = "Generating Report",
                        Detail = $"Generating '{item.Title}'...",
                        Duration = 3000
                    });
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Report Execution Error",
                        Detail = ex.Message,
                        Duration = 5000
                    });
                }
                return;
            }

            if (item.RequiresFilterModal)
            {
                ActiveReportItem = item;
                CriteriaModalTitle = $"{item.Title} Criteria";
                IsCriteriaModalOpen = true;
                StateHasChanged();
                return;
            }

            try
            {
                var request = new ReportRequest
                {
                    ReportName = item.ReportName,
                    SelectionFormula = item.SelectionFormula
                };

                await ReportNavigationService.PrintReportAsync(request);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Info,
                    Summary = "Generating Report",
                    Detail = $"Generating '{item.Title}'...",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Execution Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task HandleCriteriaModalSubmitAsync(EmpIdAndDateRangeSelection selection)
        {
            if (ActiveReportItem == null) return;

            try
            {
                ReportRequest request;

                if (ActiveReportItem.Id == "attendance-status")
                {
                    var dtFrom = new DateTime(selection.DateFrom.Year, selection.DateFrom.Month, 1);
                    var dtTo = new DateTime(selection.DateFrom.Year, selection.DateFrom.Month, DateTime.DaysInMonth(selection.DateFrom.Year, selection.DateFrom.Month));

                    request = new ReportRequest
                    {
                        ReportName = ActiveReportItem.ReportName,
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", dtFrom },
                            { "@DTTo", dtTo },
                            { "@DeptID", selection.DeptID }
                        },
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "DateRange", $"'{dtFrom:dd-MMM-yyyy} to {dtTo:dd-MMM-yyyy}'" }
                        }
                    };
                }
                else if (ActiveReportItem.Id == "employee-wise-leaves")
                {
                    string strSelection = $"{{Employees.TempDept}}=False AND {{Leaves.DT}} in Date({selection.DateFrom.Year}, {selection.DateFrom.Month}, {selection.DateFrom.Day}) to Date({selection.DateTo.Year}, {selection.DateTo.Month}, {selection.DateTo.Day})";
                    if (selection.EmpID != "0")
                    {
                        strSelection += $" AND {{Employees.EmpID}}='{selection.EmpID}'";
                    }
                    else if (selection.DeptID != "0")
                    {
                        strSelection += $" AND {{Employees.DeptID}}='{selection.DeptID}'";
                    }

                    request = new ReportRequest
                    {
                        ReportName = ActiveReportItem.ReportName,
                        SelectionFormula = strSelection,
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "MainHeading", $"'{selection.DateFrom:dd-MMM-yyyy} to {selection.DateTo:dd-MMM-yyyy}'" }
                        }
                    };
                }
                else if (ActiveReportItem.Id == "missing-outtime")
                {
                    string strSelection = $"{{EmpTimes.DT}} in Date({selection.DateFrom.Year}, {selection.DateFrom.Month}, {selection.DateFrom.Day}) to Date({selection.DateTo.Year}, {selection.DateTo.Month}, {selection.DateTo.Day})";
                    if (selection.EmpID != "0")
                    {
                        strSelection += $" AND {{Employees.EmpID}}='{selection.EmpID}'";
                    }
                    else if (selection.DeptID != "0")
                    {
                        strSelection += $" AND {{Employees.DeptID}}='{selection.DeptID}'";
                    }

                    request = new ReportRequest
                    {
                        ReportName = ActiveReportItem.ReportName,
                        SelectionFormula = strSelection,
                        GroupSelectionFormula = "IsNull({EmpTimes.OutTime})=True OR Count({EmpTimes.EmpID}, {Employees.empid})=1",
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "MainHeading", $"'{selection.DateFrom:dd-MMM-yyyy} to {selection.DateTo:dd-MMM-yyyy}'" }
                        }
                    };
                }
                else if (ActiveReportItem.Id == "late-attendance")
                {
                    string strSelection = $"{{AttendanceSheet.DT}} in Date({selection.DateFrom.Year}, {selection.DateFrom.Month}, {selection.DateFrom.Day}) to Date({selection.DateTo.Year}, {selection.DateTo.Month}, {selection.DateTo.Day})";
                    if (selection.EmpID != "0")
                    {
                        strSelection += $" AND {{Employees.EmpID}}='{selection.EmpID}'";
                    }
                    else if (selection.DeptID != "0")
                    {
                        strSelection += $" AND {{Employees.DeptID}}='{selection.DeptID}'";
                    }

                    request = new ReportRequest
                    {
                        ReportName = ActiveReportItem.ReportName,
                        SelectionFormula = strSelection,
                        FormulaValues = new Dictionary<string, object>
                        {
                            { "MainHeading", $"'{selection.DateFrom:dd-MMM-yyyy} To {selection.DateTo:dd-MMM-yyyy}'" }
                        }
                    };
                }
                else if (ActiveReportItem.Id == "employee-targets")
                {
                    string strSelection = string.Empty;
                    if (selection.EmpID != "0")
                    {
                        strSelection = $"{{Employees_Targets_Report_SP;1.EmpID}}='{selection.EmpID}'";
                    }

                    request = new ReportRequest
                    {
                        ReportName = ActiveReportItem.ReportName,
                        SelectionFormula = strSelection,
                        Parameters = new Dictionary<string, object>
                        {
                            { "@DTFrom", selection.DateFrom },
                            { "@DTTo", selection.DateTo }
                        }
                    };
                }
                else
                {
                    string strCondition = $"{{MonthlySalaries.DT}} in Date({selection.DateFrom.Year}, {selection.DateFrom.Month}, {selection.DateFrom.Day}) to Date({selection.DateTo.Year}, {selection.DateTo.Month}, {selection.DateTo.Day})";
                    if (selection.EmpID != "0")
                    {
                        strCondition += $" AND {{VEmp.EmpID}}='{selection.EmpID}'";
                    }
                    else if (selection.DeptID != "0")
                    {
                        strCondition += $" AND {{VEmp.DeptID}}='{selection.DeptID}'";
                    }

                    request = new ReportRequest
                    {
                        ReportName = ActiveReportItem.ReportName,
                        SelectionFormula = strCondition
                    };
                }

                IsCriteriaModalOpen = false;
                StateHasChanged();

                await ReportNavigationService.PrintReportAsync(request);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Info,
                    Summary = "Generating Report",
                    Detail = $"Generating '{ActiveReportItem.Title}'...",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Execution Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private void GoBack()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/payroll/dashboard");
            }
        }
    }

    public class PayrollReportCardItem
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Category { get; set; } = "General";
        public string Icon { get; set; } = "fas fa-file-pdf";
        public string BadgeText { get; set; } = string.Empty;
        public string BadgeClass { get; set; } = "bg-secondary";
        public string ReportName { get; set; } = string.Empty;
        public string SelectionFormula { get; set; } = string.Empty;
        public string NavigationUrl { get; set; } = string.Empty;
        public bool RequiresFilterModal { get; set; } = false;
        public bool RequiresDateRangeModal { get; set; } = false;
    }
}

