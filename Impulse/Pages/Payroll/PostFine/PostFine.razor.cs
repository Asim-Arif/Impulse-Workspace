using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.PostFine
{
    public partial class PostFine : ComponentBase
    {
        // ── Injected Services ─────────────────────────────────────────────────
        [Inject] private IPostFineService        _postFineService     { get; set; } = null!;
        [Inject] private IEmployeeService        _employeeService     { get; set; } = null!;
        [Inject] private INotificationService    _notificationService { get; set; } = null!;
        [Inject] private NavigationManager       _navManager          { get; set; } = null!;

        // ── Query Parameters ──────────────────────────────────────────────────
        [Parameter][SupplyParameterFromQuery] public string? returnUrl   { get; set; }
        [Parameter][SupplyParameterFromQuery] public bool   contractors { get; set; } = false;

        // ── UI State ──────────────────────────────────────────────────────────
        private bool IsLoading    { get; set; } = true;
        private bool IsSaving     { get; set; } = false;
        private bool IsContractor { get; set; } = false;

        // ── Lookups & Selections ──────────────────────────────────────────────
        private List<DepartmentListItemModel> AllDepartments     { get; set; } = new();
        private List<EmployeeListItemModel>   AllEmployees       { get; set; } = new();
        private DepartmentListItemModel?      SelectedDepartment { get; set; }
        private EmployeeListItemModel?        SelectedEmployee   { get; set; }

        // ── Form Inputs ───────────────────────────────────────────────────────
        private DateTime FineDate    { get; set; } = DateTime.Today;
        private decimal  Amount      { get; set; } = 0;
        private string   Description { get; set; } = string.Empty;

        // ── Lifecycle ─────────────────────────────────────────────────────────

        protected override async Task OnInitializedAsync()
        {
            IsContractor = contractors;
            IsLoading    = true;

            try
            {
                AllDepartments = await _employeeService.GetDepartmentsAsync(false);
                AllEmployees   = await _employeeService.GetEmployeesAsync("0", false, IsContractor);
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Load Error", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ── Event Handlers ────────────────────────────────────────────────────

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            SelectedEmployee   = null;

            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await _employeeService.GetEmployeesAsync(deptId, false, IsContractor);
        }

        private void OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
        }

        private void OnFineDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime d))
            {
                FineDate = d;
            }
        }

        // ── Typeahead Search ──────────────────────────────────────────────────

        private Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string search)
        {
            var results = AllDepartments
                .Where(d => string.IsNullOrWhiteSpace(search) ||
                            d.Name.Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        private Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string search)
        {
            var results = AllEmployees
                .Where(e => string.IsNullOrWhiteSpace(search) ||
                            e.EmpID.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            e.Name.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            (e.FName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            (e.Designation ?? "").Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        // ── Save ──────────────────────────────────────────────────────────────

        private async Task SaveAsync()
        {
            if (IsSaving) return;

            if (SelectedEmployee == null)
            {
                _notificationService.ShowWarning("Validation", "Please select an employee.");
                return;
            }
            if (Amount <= 0)
            {
                _notificationService.ShowWarning("Validation", "Please enter a valid fine amount greater than zero.");
                return;
            }

            var dto = new PostFineDto
            {
                EmpID        = SelectedEmployee.EmpID,
                FineDate     = FineDate,
                Amount       = Amount,
                Description  = Description,
                IsContractor = IsContractor
            };

            IsSaving = true;
            try
            {
                await _postFineService.SaveFineAsync(dto);

                _notificationService.ShowSuccess(
                    "Fine Posted",
                    $"Fine of {Amount:N2} posted successfully for [{SelectedEmployee.Name}].");

                await ResetForm();
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Save Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        // ── Reset Form ────────────────────────────────────────────────────────

        private async Task ResetForm()
        {
            SelectedEmployee = null;
            Amount           = 0;
            Description      = string.Empty;
            FineDate         = DateTime.Today;
            await Task.CompletedTask;
            StateHasChanged();
        }

        // ── Navigation ────────────────────────────────────────────────────────

        private void NavigateBack()
        {
            if (!string.IsNullOrEmpty(returnUrl))
                _navManager.NavigateTo(returnUrl);
            else
                _navManager.NavigateTo("/payroll");
        }
    }
}
