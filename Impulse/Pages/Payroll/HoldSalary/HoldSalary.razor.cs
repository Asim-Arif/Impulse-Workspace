using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.WebUtilities;
using Impulse.Services;

namespace Impulse.Pages.Payroll.HoldSalary
{
    public partial class HoldSalary : ComponentBase
    {
        [Inject] private IHoldSalaryService HoldSalaryService { get; set; } = default!;
        [Inject] private IEmployeeService EmployeeService { get; set; } = default!;
        [Inject] private INotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        private string? ReturnUrl { get; set; }

        // Filter / Selection state
        private List<DepartmentListItemModel> Departments { get; set; } = new();
        private List<EmployeeListItemModel> AllEmployees { get; set; } = new();
        private List<EmployeeListItemModel> FilteredEmployees { get; set; } = new();

        private DepartmentListItemModel? SelectedDepartment { get; set; }
        private EmployeeListItemModel? SelectedEmployee { get; set; }

        private int SelectedYear { get; set; } = DateTime.Today.Year;
        private int SelectedMonth { get; set; } = DateTime.Today.Month;

        // Form state
        private HoldSalaryDto Dto { get; set; } = new();
        private HoldSalarySummaryModel Summary { get; set; } = new();
        private List<HoldSalaryListItemModel> HoldSalariesList { get; set; } = new();

        private bool IsLoading { get; set; } = true;
        private bool IsSaving { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            var uri = NavigationManager.ToAbsoluteUri(NavigationManager.Uri);
            if (QueryHelpers.ParseQuery(uri.Query).TryGetValue("returnUrl", out var returnUrl))
            {
                ReturnUrl = returnUrl;
            }

            Dto.Year = SelectedYear;
            Dto.Month = SelectedMonth;

            await LoadInitialDataAsync();
        }

        private async Task LoadInitialDataAsync()
        {
            IsLoading = true;
            try
            {
                var deptsTask = EmployeeService.GetDepartmentsAsync(false);
                var empsTask = EmployeeService.GetEmployeesAsync("0", false, false);

                await Task.WhenAll(deptsTask, empsTask);

                Departments = deptsTask.Result?.ToList() ?? new List<DepartmentListItemModel>();
                AllEmployees = empsTask.Result?.ToList() ?? new List<EmployeeListItemModel>();
                FilteredEmployees = AllEmployees.ToList();

                await RefreshGridAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Load Error", $"Failed to initialize form: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task RefreshGridAsync()
        {
            try
            {
                string? deptId = SelectedDepartment?.DeptID;
                HoldSalariesList = await HoldSalaryService.GetHoldSalariesAsync(deptId, SelectedYear, SelectedMonth);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to load existing hold salaries: {ex.Message}");
            }
        }

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            SelectedEmployee = null;
            Summary = new HoldSalarySummaryModel();

            if (dept == null || dept.DeptID == "0")
            {
                FilteredEmployees = AllEmployees.ToList();
            }
            else
            {
                FilteredEmployees = AllEmployees.Where(e => e.DeptID == dept.DeptID).ToList();
            }

            await RefreshGridAsync();
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            if (emp != null)
            {
                Dto.EmpID = emp.EmpID;
                await LoadEmployeeSummaryAsync();
            }
            else
            {
                Dto.EmpID = string.Empty;
                Summary = new HoldSalarySummaryModel();
                Dto.HoldAmt = 0;
                Dto.Description = string.Empty;
            }
        }

        private async Task LoadEmployeeSummaryAsync()
        {
            if (SelectedEmployee == null) return;

            try
            {
                Summary = await HoldSalaryService.GetEmployeeHoldSalaryInfoAsync(SelectedEmployee.EmpID, SelectedYear, SelectedMonth);
                Dto.HoldAmt = Summary.ExistingHoldAmt;
                Dto.Description = Summary.ExistingDescription;
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to load employee balance info: {ex.Message}");
            }
        }

        private async Task OnMonthOrYearChanged()
        {
            Dto.Year = SelectedYear;
            Dto.Month = SelectedMonth;

            if (SelectedEmployee != null)
            {
                await LoadEmployeeSummaryAsync();
            }
            await RefreshGridAsync();
        }

        private async Task SaveHoldSalaryAsync()
        {
            if (SelectedEmployee == null)
            {
                NotificationService.ShowWarning("Validation", "Please select an employee.");
                return;
            }

            if (Dto.HoldAmt <= 0)
            {
                NotificationService.ShowWarning("Validation", "Hold amount must be greater than zero.");
                return;
            }

            if (Dto.HoldAmt > Summary.AvailableBalance)
            {
                NotificationService.ShowWarning("Validation", $"Hold amount ({Dto.HoldAmt:N0}) cannot exceed the available salary balance ({Summary.AvailableBalance:N0}).");
                return;
            }

            IsSaving = true;
            try
            {
                Dto.EmpID = SelectedEmployee.EmpID;
                Dto.Year = SelectedYear;
                Dto.Month = SelectedMonth;

                bool saved = await HoldSalaryService.SaveHoldSalaryAsync(Dto);
                if (saved)
                {
                    NotificationService.ShowSuccess("Saved", $"Hold salary of {Dto.HoldAmt:N0} recorded for [{SelectedEmployee.EmpID}] {SelectedEmployee.Name}.");
                    await LoadEmployeeSummaryAsync();
                    await RefreshGridAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        private async Task DeleteHoldSalaryAsync(HoldSalaryListItemModel item)
        {
            try
            {
                bool deleted = await HoldSalaryService.DeleteHoldSalaryAsync(item.EntryID);
                if (deleted)
                {
                    NotificationService.ShowSuccess("Removed", $"Hold salary for [{item.EmpID}] {item.EmpName} removed.");
                    if (SelectedEmployee?.EmpID == item.EmpID)
                    {
                        await LoadEmployeeSummaryAsync();
                    }
                    await RefreshGridAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Remove Failed", ex.Message);
            }
        }

        private void ResetForm()
        {
            SelectedEmployee = null;
            Dto.EmpID = string.Empty;
            Dto.HoldAmt = 0;
            Dto.Description = string.Empty;
            Summary = new HoldSalarySummaryModel();
        }

        private void GoBack()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/payroll");
            }
        }

        // ── BlazoredTypeahead Search Helpers ─────────────────────────────────────
        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            return await Task.FromResult(Departments.Where(d =>
                string.IsNullOrWhiteSpace(searchText) ||
                (d.DeptID?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (d.Name?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false)
            ));
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            return await Task.FromResult(FilteredEmployees.Where(emp =>
                string.IsNullOrWhiteSpace(searchText) ||
                (emp.EmpID?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (emp.Name?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (emp.FName?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (emp.Designation?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false)
            ).Take(1000));
        }
    }
}
