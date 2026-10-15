using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services.Payroll;

namespace Impulse.Components.Payroll.NewEmployeeTabs
{
    public partial class CopyTargetsModal
    {
        [Parameter]
        public bool IsOpen { get; set; }

        [Parameter]
        public EventCallback<bool> IsOpenChanged { get; set; }

        [Parameter]
        public string? CurrentEmpId { get; set; }

        [Parameter]
        public List<ProcessLookupModel> Processes { get; set; } = new();

        [Parameter]
        public List<GroupLookupModel> Groups { get; set; } = new();

        [Parameter]
        public EventCallback<(List<EmpDailyTargetDto> Targets, bool ReplaceExisting)> OnTargetsCopied { get; set; }

        [Inject]
        public IEmployeeService EmployeeService { get; set; } = default!;

        [Inject]
        public Radzen.NotificationService NotificationService { get; set; } = default!;

        protected List<DepartmentListItemModel> Departments { get; set; } = new();
        protected List<EmployeeListItemModel> AllEmployees { get; set; } = new();
        protected string SelectedDeptId { get; set; } = "0";
        protected EmployeeListItemModel? SelectedEmployee { get; set; }
        protected List<EmpDailyTargetDto> SourceTargets { get; set; } = new();
        protected bool IsLoadingTargets { get; set; } = false;
        protected bool ReplaceExisting { get; set; } = true;

        private bool _wasOpen = false;

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen && !_wasOpen)
            {
                _wasOpen = true;
                await InitializeDataAsync();
            }
            else if (!IsOpen)
            {
                _wasOpen = false;
            }
        }

        private async Task InitializeDataAsync()
        {
            SelectedDeptId = "0";
            SelectedEmployee = null;
            SourceTargets.Clear();
            ReplaceExisting = true;

            try
            {
                Departments = await EmployeeService.GetDepartmentsAsync(false);
                await LoadEmployeesAsync("0");
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Unable to load employees: {ex.Message}",
                    Duration = 4000
                });
            }
        }

        private async Task LoadEmployeesAsync(string deptId)
        {
            var rawList = await EmployeeService.GetEmployeesAsync(deptId, false, false);
            // Exclude current employee from copy list
            AllEmployees = rawList
                .Where(e => string.IsNullOrWhiteSpace(CurrentEmpId) || !string.Equals(e.EmpID, CurrentEmpId, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        protected async Task OnDepartmentFilterChanged(ChangeEventArgs e)
        {
            SelectedDeptId = e?.Value?.ToString() ?? "0";
            SelectedEmployee = null;
            SourceTargets.Clear();
            await LoadEmployeesAsync(SelectedDeptId);
        }

        protected Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Task.FromResult<IEnumerable<EmployeeListItemModel>>(AllEmployees.Take(100));
            }

            var query = AllEmployees.Where(e =>
                (!string.IsNullOrEmpty(e.EmpID) && e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.Name) && e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.FName) && e.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.Designation) && e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            ).Take(100);

            return Task.FromResult<IEnumerable<EmployeeListItemModel>>(query);
        }

        protected async Task OnEmployeeSelected(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            SourceTargets.Clear();

            if (emp != null && !string.IsNullOrEmpty(emp.EmpID))
            {
                IsLoadingTargets = true;
                try
                {
                    var empDetail = await EmployeeService.GetEmployeeByIdAsync(emp.EmpID);
                    SourceTargets = empDetail?.DailyTargets ?? new List<EmpDailyTargetDto>();
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Error",
                        Detail = $"Failed to load employee targets: {ex.Message}",
                        Duration = 4000
                    });
                }
                finally
                {
                    IsLoadingTargets = false;
                }
            }
        }

        protected async Task ConfirmCopy()
        {
            if (SelectedEmployee == null || SourceTargets.Count == 0) return;

            // Deep clone targets, resetting EntryID = 0
            var cloned = SourceTargets.Select(t => new EmpDailyTargetDto
            {
                EntryID = 0,
                ProcessID = t.ProcessID,
                GroupID = t.GroupID,
                Qty = t.Qty,
                OverTime = t.OverTime
            }).ToList();

            if (OnTargetsCopied.HasDelegate)
            {
                await OnTargetsCopied.InvokeAsync((cloned, ReplaceExisting));
            }

            await CloseModal();
        }

        protected async Task CloseModal()
        {
            IsOpen = false;
            if (IsOpenChanged.HasDelegate)
            {
                await IsOpenChanged.InvokeAsync(false);
            }
        }
    }
}
