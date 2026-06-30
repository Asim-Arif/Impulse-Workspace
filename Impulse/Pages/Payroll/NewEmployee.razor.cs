using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services.Payroll;

namespace Impulse.Pages.Payroll
{
    public partial class NewEmployee
    {
        [Parameter]
        public string? EmployeeId { get; set; }

        [Inject]
        public IEmployeeService EmployeeService { get; set; } = default!;

        [Inject]
        public IDepartmentService DepartmentService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        public Radzen.NotificationService NotificationService { get; set; } = default!;

        public EmployeeDto Employee { get; set; } = new EmployeeDto();
        public List<DepartmentModel> Departments { get; set; } = new List<DepartmentModel>();
        public List<string> Designations { get; set; } = new List<string>();
        public List<ProcessLookupModel> Processes { get; set; } = new List<ProcessLookupModel>();
        public List<GroupLookupModel> Groups { get; set; } = new List<GroupLookupModel>();

        public string ActiveTab { get; set; } = "PersonalInfo";
        public bool IsSaving { get; set; } = false;

        public DepartmentModel? SelectedDepartment { get; set; }

        public bool ExemptSettingsHelper
        {
            get => Employee.ExemptSettings ?? false;
            set
            {
                Employee.ExemptSettings = value;
                if (!value && ActiveTab == "PayrollSettings")
                {
                    ActiveTab = "PersonalInfo";
                }
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();

            if (!string.IsNullOrEmpty(EmployeeId))
            {
                var emp = await EmployeeService.GetEmployeeByIdAsync(EmployeeId);
                if (emp != null)
                {
                    Employee = emp;
                    SelectedDepartment = Departments.FirstOrDefault(d => d.DeptID == Employee.DeptID);
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Error",
                        Detail = $"Employee not found.",
                        Duration = 4000
                    });
                }
            }
            else
            {
                // Default values
                Employee.JoinDate = DateTime.Today;
                Employee.Active = true;
                Employee.ExemptSettings = false;
            }
        }

        private async Task LoadLookups()
        {
            try
            {
                Departments = await DepartmentService.GetDepartmentsListAsync(true);
                Designations = new List<string> { "Manager", "Supervisor", "Staff", "Worker" };
                Processes = await EmployeeService.GetProcessesAsync();
                Groups = await EmployeeService.GetGroupsAsync();

                if (!string.IsNullOrEmpty(Employee.DeptID))
                {
                    SelectedDepartment = Departments.FirstOrDefault(d => d.DeptID == Employee.DeptID);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load lookup data: {ex.Message}",
                    Duration = 4000
                });
            }
        }

        private async Task OnDepartmentSelected(DepartmentModel? dept)
        {
            await InvokeAsync(async () =>
            {
                SelectedDepartment = dept;
                if (dept != null)
                {
                    Employee.DeptID = dept.DeptID;
                    await HandleDepartmentChanged(dept.DeptID);
                }
                else
                {
                    Employee.DeptID = string.Empty;
                }
                StateHasChanged();
            });
        }

        private async Task<IEnumerable<DepartmentModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Departments;
            return await Task.FromResult(Departments.Where(d => d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task HandleDepartmentChanged(string deptId)
        {
            if (!string.IsNullOrEmpty(deptId) && string.IsNullOrEmpty(Employee.EmpID))
            {
                try
                {
                    var newEmpId = await EmployeeService.GetNextEmpIDAsync(deptId);
                    await InvokeAsync(() =>
                    {
                        Employee.EmpID = newEmpId;
                        StateHasChanged();
                    });
                }
                catch (Exception ex)
                {
                    await InvokeAsync(() =>
                    {
                        NotificationService.Notify(new Radzen.NotificationMessage
                        {
                            Severity = Radzen.NotificationSeverity.Error,
                            Summary = "Error",
                            Detail = $"Failed to generate Employee ID: {ex.Message}",
                            Duration = 4000
                        });
                        StateHasChanged();
                    });
                }
            }
        }

        private void ChangeTab(string tabName)
        {
            ActiveTab = tabName;
        }

        private async Task HandleValidSubmit()
        {
            if (string.IsNullOrWhiteSpace(Employee.EmpID))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Employee ID is required.",
                    Duration = 4000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(Employee.Name))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Employee Name is required.",
                    Duration = 4000
                });
                return;
            }

            IsSaving = true;
            try
            {
                bool isAdd = string.IsNullOrEmpty(EmployeeId);
                await EmployeeService.SaveEmployeeAsync(Employee, isAdd);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Success",
                    Detail = "Employee saved successfully.",
                    Duration = 4000
                });

                NavigationManager.NavigateTo("/payroll/employeelist");
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to save employee: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo("/payroll/employeelist");
        }
    }
}
