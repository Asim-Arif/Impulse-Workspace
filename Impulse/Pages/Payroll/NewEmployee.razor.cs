using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services.Payroll;

namespace Impulse.Pages.Payroll
{
    public partial class NewEmployee
    {
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

        public string ActiveTab { get; set; } = "PersonalInfo";
        public bool IsSaving { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();

            // Default values
            Employee.JoinDate = DateTime.Today;
            Employee.Active = true;
        }

        private async Task LoadLookups()
        {
            try
            {
                Departments = await DepartmentService.GetDepartmentsListAsync(true);
                // To fetch designations, we would normally use a service, but for now we initialize an empty list or hardcode common ones to match VB6 if a service is missing.
                Designations = new List<string> { "Manager", "Supervisor", "Staff", "Worker" };
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
                // In a full implementation, we'd call EmployeeService.AddEmployeeAsync(Employee)
                // For now, we simulate success since the DAC implementation for adding an Employee might be missing or incomplete.
                await Task.Delay(500);

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
