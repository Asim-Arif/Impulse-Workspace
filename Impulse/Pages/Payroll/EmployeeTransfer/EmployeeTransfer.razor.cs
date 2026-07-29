using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.EmployeeTransfer
{
    public partial class EmployeeTransfer : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        private bool IsSaving = false;
        private bool IsLoadingHistory = false;
        private string CurrentUserName = "System";

        private DateTime TransferDate { get; set; } = DateTime.Today;
        private string Description { get; set; } = string.Empty;

        private List<DepartmentListItemModel> SourceDepartments = new();
        private DepartmentListItemModel _selectedSourceDepartment;
        public DepartmentListItemModel SelectedSourceDepartment
        {
            get => _selectedSourceDepartment;
            set
            {
                if (_selectedSourceDepartment != value)
                {
                    _selectedSourceDepartment = value;
                    _ = OnSourceDepartmentChanged();
                }
            }
        }

        private List<EmployeeListItemModel> Employees = new();
        private EmployeeListItemModel SelectedEmployee { get; set; }

        private List<DepartmentListItemModel> TargetDepartments = new();
        private DepartmentListItemModel SelectedTargetDepartment { get; set; }

        private List<EmployeeTransferHistoryRow> RecentTransfers = new();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                SourceDepartments = await EmployeeTransferService.GetSourceDepartmentsAsync();
                await LoadRecentTransfersAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        private async Task OnSourceDepartmentChanged()
        {
            SelectedEmployee = null;
            SelectedTargetDepartment = null;
            Employees.Clear();
            TargetDepartments.Clear();

            if (SelectedSourceDepartment == null)
            {
                await InvokeAsync(StateHasChanged);
                return;
            }

            try
            {
                string deptId = SelectedSourceDepartment.DeptID;
                Employees = await EmployeeTransferService.GetEmployeesByDepartmentAsync(deptId);
                TargetDepartments = await EmployeeTransferService.GetTargetDepartmentsAsync(deptId);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Department Data", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task<IEnumerable<DepartmentListItemModel>> SearchSourceDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return SourceDepartments;

            return await Task.FromResult(SourceDepartments.Where(x => x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Employees;

            return await Task.FromResult(Employees.Where(x =>
                x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                x.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                (x.FName != null && x.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (x.Designation != null && x.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            ).ToList());
        }

        private async Task<IEnumerable<DepartmentListItemModel>> SearchTargetDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return TargetDepartments;

            return await Task.FromResult(TargetDepartments.Where(x => x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task SaveTransferAsync(bool closeAfterSave)
        {
            if (SelectedSourceDepartment == null)
            {
                NotificationService.ShowWarning("Validation Warning", "Please select a source department.");
                return;
            }

            if (SelectedEmployee == null)
            {
                NotificationService.ShowWarning("Validation Warning", "Please select an employee to transfer.");
                return;
            }

            if (SelectedTargetDepartment == null)
            {
                NotificationService.ShowWarning("Validation Warning", "Please select a target department for transfer.");
                return;
            }

            IsSaving = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                var dto = new EmployeeTransferDto
                {
                    EmpID = SelectedEmployee.EmpID,
                    TransferFromDeptID = SelectedSourceDepartment.DeptID,
                    TransferToDeptID = SelectedTargetDepartment.DeptID,
                    TransferDT = TransferDate,
                    EntryDescription = Description,
                    UserName = CurrentUserName
                };

                bool success = await EmployeeTransferService.SaveTransferAsync(dto);
                if (success)
                {
                    NotificationService.ShowSuccess("Transfer Saved", $"Employee '{SelectedEmployee.Name}' transferred successfully to {SelectedTargetDepartment.Name}.");
                    await LoadRecentTransfersAsync();

                    if (closeAfterSave)
                    {
                        NavigateBack();
                    }
                    else
                    {
                        ResetForm();
                    }
                }
                else
                {
                    NotificationService.ShowError("Error", "Failed to process employee transfer.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Transfer Exception", ex.Message);
            }
            finally
            {
                IsSaving = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private void ResetForm()
        {
            _selectedSourceDepartment = null;
            SelectedEmployee = null;
            SelectedTargetDepartment = null;
            Employees.Clear();
            TargetDepartments.Clear();
            TransferDate = DateTime.Today;
            Description = string.Empty;
        }

        private async Task LoadRecentTransfersAsync()
        {
            IsLoadingHistory = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                RecentTransfers = await EmployeeTransferService.GetRecentTransfersAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", "Failed to load recent transfer history: " + ex.Message);
            }
            finally
            {
                IsLoadingHistory = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/payroll");
            }
        }
    }
}
