using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.AttendanceLedger
{
    public partial class AttendanceLedger : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        private bool IsLoading = false;
        private bool IsPrinting = false;
        private string CompanyName = string.Empty;

        private static readonly string[] MonthNames =
            { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

        private int SelectedMonth { get; set; } = DateTime.Today.Month;
        private int SelectedYear { get; set; } = DateTime.Today.Year;
        private List<int> Years = new();

        private List<DepartmentListItemModel> Departments = new();
        private List<EmployeeListItemModel> Employees = new();

        private DepartmentListItemModel _selectedDepartment;
        private DepartmentListItemModel SelectedDepartment
        {
            get => _selectedDepartment;
            set
            {
                if (_selectedDepartment != value)
                {
                    _selectedDepartment = value;
                    _ = OnFilterChangedAsync();
                }
            }
        }

        private EmployeeListItemModel _selectedEmployee;
        private EmployeeListItemModel SelectedEmployee
        {
            get => _selectedEmployee;
            set
            {
                if (_selectedEmployee != value)
                {
                    _selectedEmployee = value;
                    _ = LoadLedgerAsync();
                }
            }
        }

        private List<AttendanceLedgerDayRow> DayRows = new();

        // Totals (mirror the legacy print formulas)
        private double TotalHrs { get; set; }
        private double TotalPayableHrs { get; set; }
        private double TotalOTHrs { get; set; }
        private double TotalLateHrs { get; set; }
        private int OffDayCount { get; set; }
        private int PresentCount { get; set; }
        private int AbsentCount { get; set; }
        private int LeaveCount { get; set; }
        private int HalfLeaveCount { get; set; }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                int currentYear = DateTime.Today.Year;
                for (int y = 2002; y <= currentYear; y++) Years.Add(y);

                CompanyName = await LedgerService.GetCompanyNameAsync();
                Departments = await LedgerService.GetDepartmentsAsync();
                Employees = await LedgerService.GetEmployeesAsync("0", FirstDayOfMonth, LastDayOfMonth);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        private DateTime FirstDayOfMonth => new DateTime(SelectedYear, SelectedMonth, 1);
        private DateTime LastDayOfMonth => FirstDayOfMonth.AddMonths(1).AddDays(-1);

        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Departments;

            return await Task.FromResult(Departments.Where(x => x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Employees;

            return await Task.FromResult(Employees.Where(x =>
                x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                x.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                (!string.IsNullOrWhiteSpace(x.FName) && x.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrWhiteSpace(x.Designation) && x.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))).ToList());
        }

        private async Task OnMonthChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e?.Value?.ToString(), out int m) && m >= 1 && m <= 12)
            {
                SelectedMonth = m;
                await OnMonthYearChangedAsync();
            }
        }

        private async Task OnYearChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e?.Value?.ToString(), out int y))
            {
                SelectedYear = y;
                await OnMonthYearChangedAsync();
            }
        }

        private static string GetStatusBadgeClass(string status) => status switch
        {
            "P" => "bg-success",
            "A" => "bg-danger",
            "HL" => "bg-warning text-dark",
            "L" => "bg-info text-dark",
            _ => "bg-secondary"
        };

        private async Task OnMonthYearChangedAsync()
        {
            await OnFilterChangedAsync();
        }

        private async Task OnFilterChangedAsync()
        {
            _selectedEmployee = null;
            DayRows = new();
            ResetTotals();

            try
            {
                Employees = await LedgerService.GetEmployeesAsync(SelectedDepartment?.DeptID ?? "0", FirstDayOfMonth, LastDayOfMonth);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private void ResetTotals()
        {
            TotalHrs = 0; TotalPayableHrs = 0; TotalOTHrs = 0; TotalLateHrs = 0;
            OffDayCount = 0; PresentCount = 0; AbsentCount = 0; LeaveCount = 0; HalfLeaveCount = 0;
        }

        /// <summary>C# port of the legacy Refresh2ShiftLedger() routine.</summary>
        private async Task LoadLedgerAsync()
        {
            DayRows = new();
            ResetTotals();

            if (SelectedEmployee == null || string.IsNullOrWhiteSpace(SelectedEmployee.EmpID))
                return;

            IsLoading = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                var filter = new AttendanceLedgerFilterModel
                {
                    DeptID = SelectedDepartment?.DeptID ?? "0",
                    EmpID = SelectedEmployee.EmpID,
                    Month = SelectedMonth,
                    Year = SelectedYear
                };

                var holidays = await LedgerService.GetHolidaysAsync(SelectedMonth, SelectedYear);
                var entries = await LedgerService.GetLedgerEntriesAsync(filter);
                var times = await LedgerService.GetTimeEntriesAsync(filter);
                var shortLeaves = await LedgerService.GetShortLeavesAsync(filter);

                var holidayByDay = holidays.GroupBy(h => h.Day).ToDictionary(g => g.Key, g => g.First().Description);
                var entryByDay = entries.GroupBy(e => e.AttDay).ToDictionary(g => g.Key, g => g.First());
                var timeByDate = times.GroupBy(t => t.DT.Date).ToDictionary(g => g.Key, g => g.First());

                int daysInMonth = DateTime.DaysInMonth(SelectedYear, SelectedMonth);

                for (int day = 1; day <= daysInMonth; day++)
                {
                    var date = new DateTime(SelectedYear, SelectedMonth, day);
                    bool isSunday = date.DayOfWeek == DayOfWeek.Sunday;
                    bool isHoliday = holidayByDay.TryGetValue(day, out string holidayDesc);

                    var row = new AttendanceLedgerDayRow
                    {
                        Date = date,
                        DayName = date.ToString("dddd", CultureInfo.InvariantCulture) +
                                  (isHoliday && !string.IsNullOrWhiteSpace(holidayDesc) ? $"  {{{holidayDesc}}}" : string.Empty),
                        IsSunday = isSunday,
                        IsHoliday = isHoliday
                    };

                    if (isSunday || isHoliday) OffDayCount++;

                    if (entryByDay.TryGetValue(day, out var entry))
                    {
                        row.HasRecord = true;
                        row.LeaveType = ResolveLeaveType(entry);
                        row.OTHrsOther = entry.OTHrsOther ?? 0;

                        if (isSunday || isHoliday)
                        {
                            row.Status = "P";
                        }
                        else
                        {
                            row.Status = entry.Attendance switch
                            {
                                0 => "A",
                                1 => "P",
                                4 => "HL",
                                _ => "L"
                            };
                        }

                        switch (row.Status)
                        {
                            case "P" when !row.IsOffDay: PresentCount++; break;
                            case "A": AbsentCount++; break;
                            case "L": LeaveCount++; break;
                            case "HL": HalfLeaveCount++; break;
                        }

                        // Legacy: absent / leave weekdays carry no punch times
                        bool noTimes = (row.Status == "A" || row.Status == "L") && !isSunday;
                        if (!noTimes)
                        {
                            row.ShowTimes = true;

                            if (timeByDate.TryGetValue(date.Date, out var t))
                            {
                                row.FirstInTime = t.FirstInTime;
                                row.FirstOutTime = t.FirstOutTime;
                                row.SecondInTime = t.SecondInTime;
                                row.SecondOutTime = t.SecondOutTime;
                                row.OTInTime = t.OTInTime;
                                row.OTOutTime = t.OTOutTime;
                                row.Hrs = t.Hrs;
                                row.PayableHrs = t.PayableHrs;
                                row.OTHrs = t.OTHrs;
                                row.LateHours = t.LateHours;

                                TotalHrs += t.Hrs;
                                TotalPayableHrs += Math.Round(Math.Min(t.PayableHrs, 8), 2);
                                TotalOTHrs += t.OTHrs;
                                TotalLateHrs += t.LateHours;
                            }

                            if (shortLeaves.TryGetValue(date.Date, out var sl))
                                row.ShortLeaves = sl;
                        }
                    }

                    DayRows.Add(row);
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading attendance ledger", ex.Message);
            }
            finally
            {
                IsLoading = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private static string ResolveLeaveType(AttendanceLedgerEntryModel entry)
        {
            if (entry.CL != 0) return "( Casual Leave )";
            if (entry.SL != 0) return "( Sick Leave )";
            if (entry.AL != 0) return "( Annual Leave )";
            if (entry.CPL != 0) return "( Compensatory Leave )";
            if (entry.LWP != 0) return "( Leave Without Pay )";
            if (entry.ML != 0) return "( Maternity Leave )";
            if (entry.SAL != 0) return "( Special Approved Leave )";
            return string.Empty;
        }

        private async Task PrintLedgerAsync()
        {
            if (!DayRows.Any() || SelectedEmployee == null || IsPrinting) return;

            IsPrinting = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                // Fill the legacy PrintEmpLedger staging table required by the Crystal Report
                await LedgerService.PrepareEmpLedgerPrintAsync(DayRows, SelectedEmployee.EmpID);

                int daysInMonth = DateTime.DaysInMonth(SelectedYear, SelectedMonth);
                string empDisplay = $"{{{SelectedEmployee.EmpID}}} {SelectedEmployee.Name} - {SelectedEmployee.FName} {{{SelectedEmployee.Designation}}}";

                var req = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.EmpLedger,
                    SelectionFormula = string.Empty,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Company", $"'{CompanyName}'" },
                        { "FromTo", $"'For {MonthNames[SelectedMonth - 1]}, {SelectedYear}.'" },
                        { "EmpName", $"'{empDisplay}'" },
                        { "TotalLateHrs", $"'{AttendanceLedgerFormatter.GetHrsMin(TotalLateHrs)}'" },
                        { "TotalOverTime", $"'{AttendanceLedgerFormatter.GetHrsMin(TotalOTHrs)}'" },
                        { "TotalPresents", $"'{AttendanceLedgerFormatter.GetHrsMin(TotalHrs)}'" },
                        { "TotalPayableHrs", TotalPayableHrs },
                        { "TotalAbsents", (daysInMonth - OffDayCount - 1) * 8 },
                        { "ForSA", true }
                    }
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
            finally
            {
                IsPrinting = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        // ── Legacy right-click menu actions (targets not yet migrated) ──────
        private void ChangeToAbsentLeave(AttendanceLedgerDayRow row)
        {
            NotificationService.ShowWarning("Pending Migration",
                $"Leave / absent marking for {SelectedEmployee?.Name} on {row.Date:dd-MMM-yyyy} belongs to the legacy Leaves screen (frmLeaves), which has not been migrated yet.");
        }

        private void NewShortTermAdvance(AttendanceLedgerDayRow row)
        {
            NotificationService.ShowWarning("Pending Migration",
                "Short term advances (legacy NewAdvanceShort screen) have not been migrated yet.");
        }

        private void NewLongTermAdvance(AttendanceLedgerDayRow row)
        {
            NotificationService.ShowWarning("Pending Migration",
                "Long term advances (legacy NewAdvanceLong screen) have not been migrated yet.");
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
                NavigationManager.NavigateTo(ReturnUrl);
            else
                NavigationManager.NavigateTo("/payroll");
        }
    }
}
