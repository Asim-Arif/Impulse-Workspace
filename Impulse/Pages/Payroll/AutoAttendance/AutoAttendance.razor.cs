using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.AutoAttendance
{
    public partial class AutoAttendance : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<DepartmentListItemModel> DepartmentList = new();
        private List<EmployeeListItemModel> ActiveEmployees = new();

        private DepartmentListItemModel SelectedDepartment = new() { DeptID = "0", Name = "<All Departments>" };
        private EmployeeListItemModel SelectedEmployee = new() { EmpID = "0", Name = "<All Active Employees>" };

        private DateTime DateFrom { get; set; } = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        private DateTime DateTo { get; set; } = DateTime.Today;

        private bool IsProcessing = false;

        protected override async Task OnInitializedAsync()
        {
            await LoadMasterDataAsync();
        }

        private async Task LoadMasterDataAsync()
        {
            try
            {
                var depts = await OverTimeAuthService.GetDepartmentsAsync();
                DepartmentList = new List<DepartmentListItemModel>
                {
                    new DepartmentListItemModel { DeptID = "0", Name = "<All Departments>" }
                };
                if (depts != null)
                {
                    DepartmentList.AddRange(depts);
                }

                var emps = await ManualAttendanceService.GetActiveEmployeesAsync();
                ActiveEmployees = emps ?? new List<EmployeeListItemModel>();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error Loading Data", ex.Message);
            }
        }

        private Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<DepartmentListItemModel>>(DepartmentList);

            return Task.FromResult<IEnumerable<DepartmentListItemModel>>(
                DepartmentList.Where(d => d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        private Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            var filtered = ActiveEmployees.AsEnumerable();

            if (SelectedDepartment != null && SelectedDepartment.DeptID != "0")
            {
                filtered = filtered.Where(e => e.DeptID == SelectedDepartment.DeptID);
            }

            var resultList = new List<EmployeeListItemModel>
            {
                new EmployeeListItemModel { EmpID = "0", Name = "<All Active Employees>" }
            };

            if (string.IsNullOrWhiteSpace(searchText))
            {
                resultList.AddRange(filtered);
            }
            else
            {
                resultList.AddRange(filtered.Where(e =>
                    e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrEmpty(e.FName) && e.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(e.Designation) && e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
            }

            return Task.FromResult<IEnumerable<EmployeeListItemModel>>(resultList);
        }

        private void OnDepartmentChanged(DepartmentListItemModel dept)
        {
            SelectedDepartment = dept ?? new DepartmentListItemModel { DeptID = "0", Name = "<All Departments>" };
            SelectedEmployee = new EmployeeListItemModel { EmpID = "0", Name = "<All Active Employees>" };
        }

        private void OnEmployeeChanged(EmployeeListItemModel emp)
        {
            SelectedEmployee = emp ?? new EmployeeListItemModel { EmpID = "0", Name = "<All Active Employees>" };
        }

        private async Task HandleMakePresentAllDays()
        {
            if (DateFrom > DateTo)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Invalid Date Range", "'Date From' cannot be greater than 'Date To'.");
                return;
            }

            IsProcessing = true;
            StateHasChanged();

            try
            {
                string deptId = SelectedDepartment?.DeptID ?? "0";
                string empId = SelectedEmployee?.EmpID ?? "0";

                int processedCount = await TakeAttendanceService.AutoGenerateAttendanceAsync(deptId, empId, DateFrom, DateTo);

                NotificationService.Notify(NotificationSeverity.Success, "Auto Attendance Generated",
                    $"Successfully generated attendance records for {processedCount} date/employee entries.");
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Process Failed", ex.Message);
            }
            finally
            {
                IsProcessing = false;
                StateHasChanged();
            }
        }

        private async Task HandleCompleteMissingAttendance()
        {
            if (DateFrom > DateTo)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Invalid Date Range", "'Date From' cannot be greater than 'Date To'.");
                return;
            }

            IsProcessing = true;
            StateHasChanged();

            try
            {
                string deptId = SelectedDepartment?.DeptID ?? "0";
                string empId = SelectedEmployee?.EmpID ?? "0";

                int completedCount = await TakeAttendanceService.CompleteMissingAttendanceAsync(deptId, empId, DateFrom, DateTo);

                NotificationService.Notify(NotificationSeverity.Success, "Missing Attendance Completed",
                    $"Successfully completed missing punches for {completedCount} date/employee entries.");
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Process Failed", ex.Message);
            }
            finally
            {
                IsProcessing = false;
                StateHasChanged();
            }
        }

        private void GoBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                Navigation.NavigateTo(ReturnUrl);
            }
            else
            {
                Navigation.NavigateTo("/payroll-dashboard");
            }
        }
    }
}
