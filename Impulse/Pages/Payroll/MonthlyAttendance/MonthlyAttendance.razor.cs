using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.MonthlyAttendance
{
    public partial class MonthlyAttendance : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        // ── Injected Services ──────────────────────────────────────────────
        [Inject] private IMonthlyAttendanceService MonthlyAttendanceService { get; set; } = default!;
        [Inject] private IEmployeeService          EmployeeService          { get; set; } = default!;
        [Inject] private INotificationService      NotificationService      { get; set; } = default!;
        [Inject] private NavigationManager         NavigationManager        { get; set; } = default!;

        // ── State ──────────────────────────────────────────────────────────
        private bool IsLoading = false;
        private bool IsSaving  = false;

        private DepartmentListItemModel? SelectedDepartment { get; set; }
        private EmployeeListItemModel?   SelectedEmployee   { get; set; }

        private List<DepartmentListItemModel> AllDepartments = new();
        private List<EmployeeListItemModel>   AllEmployees   = new();
        private List<MonthlyAttendanceDayRow> DayRows        = new();

        private int CurrentYear  = DateTime.Today.Year;
        private int CurrentMonth = DateTime.Today.Month;
        private string SelectedMonthYear => $"{CurrentYear:D4}-{CurrentMonth:D2}";

        // ── Initialization ─────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            try
            {
                AllDepartments = await EmployeeService.GetDepartmentsAsync(false);
                AllEmployees   = await EmployeeService.GetEmployeesAsync("0", false, false);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        // ── Filters & Search ───────────────────────────────────────────────
        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AllDepartments;
            return await Task.FromResult(
                AllDepartments.Where(d => d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AllEmployees;

            return await Task.FromResult(
                AllEmployees.Where(e =>
                    e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrEmpty(e.FName) && e.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(e.Designation) && e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                )
            );
        }

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await EmployeeService.GetEmployeesAsync(deptId, false, false);
            SelectedEmployee = null;
            DayRows = new();
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            await LoadMonthGridAsync();
        }

        private async Task OnMonthYearChanged(ChangeEventArgs e)
        {
            string? val = e?.Value?.ToString();
            if (!string.IsNullOrWhiteSpace(val) && DateTime.TryParse($"{val}-01", out DateTime dt))
            {
                CurrentYear  = dt.Year;
                CurrentMonth = dt.Month;
                await LoadMonthGridAsync();
            }
        }

        private async Task LoadMonthGridAsync()
        {
            if (SelectedEmployee == null || string.IsNullOrWhiteSpace(SelectedEmployee.EmpID) || SelectedEmployee.EmpID == "0")
            {
                DayRows = new();
                return;
            }

            IsLoading = true;
            StateHasChanged();

            try
            {
                DayRows = await MonthlyAttendanceService.GetMonthlyAttendanceAsync(SelectedEmployee.EmpID, CurrentYear, CurrentMonth);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Attendance", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ── Attendance Editing ─────────────────────────────────────────────
        private void OnInTimeChanged(MonthlyAttendanceDayRow row, ChangeEventArgs e)
        {
            row.InTime = e?.Value?.ToString() ?? string.Empty;
            row.IsModified = true;
            if (!string.IsNullOrWhiteSpace(row.InTime) && row.Status == "Absent")
            {
                row.Status = "Present";
            }
        }

        private void OnOutTimeChanged(MonthlyAttendanceDayRow row, ChangeEventArgs e)
        {
            row.OutTime = e?.Value?.ToString() ?? string.Empty;
            row.IsModified = true;
        }

        private async Task ClearDateAsync(MonthlyAttendanceDayRow row)
        {
            if (SelectedEmployee == null) return;
            try
            {
                bool success = await MonthlyAttendanceService.ClearDateAttendanceAsync(SelectedEmployee.EmpID, row.Date);
                if (success)
                {
                    row.InTime = string.Empty;
                    row.OutTime = string.Empty;
                    row.Status = "Absent";
                    row.OtHours = 0;
                    row.IsModified = false;
                    NotificationService.ShowSuccess("Cleared", $"Attendance cleared for {row.Date:dd-MMM-yyyy}.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Clear Failed", ex.Message);
            }
        }

        private async Task SaveMonthAsync()
        {
            if (SelectedEmployee == null || string.IsNullOrWhiteSpace(SelectedEmployee.EmpID)) return;

            IsSaving = true;
            StateHasChanged();

            try
            {
                var input = new MonthlyAttendanceSaveDto
                {
                    EmpID = SelectedEmployee.EmpID,
                    Year  = CurrentYear,
                    Month = CurrentMonth,
                    Rows  = DayRows
                };

                bool success = await MonthlyAttendanceService.SaveMonthlyAttendanceAsync(input);
                if (success)
                {
                    NotificationService.ShowSuccess("Saved", "Monthly attendance saved successfully.");
                    await LoadMonthGridAsync();
                }
                else
                {
                    NotificationService.ShowError("Save Failed", "Could not save monthly attendance.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
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
