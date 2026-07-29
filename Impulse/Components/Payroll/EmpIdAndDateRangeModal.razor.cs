using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Components.Payroll
{
    public class EmpIdAndDateRangeSelection
    {
        public string DeptID { get; set; } = "0";
        public string EmpID { get; set; } = "0";
        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
    }

    public partial class EmpIdAndDateRangeModal : ComponentBase
    {
        [Parameter] public string Title { get; set; } = "Re-Calculate Attendance";
        [Parameter] public bool IsOpen { get; set; } = false;
        [Parameter] public EventCallback<bool> IsOpenChanged { get; set; }
        [Parameter] public EventCallback<EmpIdAndDateRangeSelection> OnSubmit { get; set; }

        private bool IsProcessing = false;

        private List<DepartmentListItemModel> Departments = new();
        private List<EmployeeListItemModel> Employees = new();

        private DepartmentListItemModel SelectedDepartment { get; set; } = new DepartmentListItemModel { DeptID = "0", Name = "<All Departments>" };
        private EmployeeListItemModel SelectedEmployee { get; set; } = new EmployeeListItemModel { EmpID = "0", Name = "<All Employees>" };

        private DateTime DateFrom { get; set; } = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        private DateTime DateTo { get; set; } = DateTime.Today;

        private static readonly DepartmentListItemModel AllDepartmentsItem = new DepartmentListItemModel { DeptID = "0", Name = "<All Departments>" };
        private static readonly EmployeeListItemModel AllEmployeesItem = new EmployeeListItemModel { EmpID = "0", Name = "<All Employees>" };

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen && (!Departments.Any() || !Employees.Any()))
            {
                await LoadDataAsync();
            }
        }

        private async Task LoadDataAsync()
        {
            try
            {
                var depts = await OverTimeAuthService.GetDepartmentsAsync();
                Departments = new List<DepartmentListItemModel> { AllDepartmentsItem };
                Departments.AddRange(depts);

                var emps = await AttendanceService.GetActiveEmployeesAsync();
                Employees = new List<EmployeeListItemModel> { AllEmployeesItem };
                Employees.AddRange(emps);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading modal lookup data: {ex.Message}");
            }
        }

        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Departments;

            return await Task.FromResult(Departments.Where(d =>
                d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)
            ).ToList());
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            var filtered = Employees.AsEnumerable();

            if (SelectedDepartment != null && SelectedDepartment.DeptID != "0")
            {
                filtered = filtered.Where(e => e.EmpID == "0" || e.DeptID == SelectedDepartment.DeptID);
            }

            if (!string.IsNullOrWhiteSpace(searchText))
            {
                filtered = filtered.Where(e =>
                    e.EmpID == "0" ||
                    e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    (e.FName != null && e.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                );
            }

            return await Task.FromResult(filtered.ToList());
        }

        private void OnDepartmentChanged(DepartmentListItemModel dept)
        {
            SelectedDepartment = dept ?? AllDepartmentsItem;
            // Reset employee selection if selected employee does not belong to selected department
            if (SelectedDepartment.DeptID != "0" && SelectedEmployee.EmpID != "0" && SelectedEmployee.DeptID != SelectedDepartment.DeptID)
            {
                SelectedEmployee = AllEmployeesItem;
            }
        }

        private void OnEmployeeChanged(EmployeeListItemModel emp)
        {
            SelectedEmployee = emp ?? AllEmployeesItem;
        }

        private void OnDateFromChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime parsed))
            {
                DateFrom = parsed.Date;
            }
        }

        private void OnDateToChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime parsed))
            {
                DateTo = parsed.Date;
            }
        }

        private async Task SubmitAsync()
        {
            IsProcessing = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                var selection = new EmpIdAndDateRangeSelection
                {
                    DeptID = SelectedDepartment?.DeptID ?? "0",
                    EmpID = SelectedEmployee?.EmpID ?? "0",
                    DateFrom = DateFrom,
                    DateTo = DateTo
                };

                await OnSubmit.InvokeAsync(selection);
                await CloseModal();
            }
            finally
            {
                IsProcessing = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task CloseModal()
        {
            IsOpen = false;
            await IsOpenChanged.InvokeAsync(false);
        }
    }
}
