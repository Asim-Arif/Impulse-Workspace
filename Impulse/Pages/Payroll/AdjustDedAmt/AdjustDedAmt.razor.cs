using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.AdjustDedAmt
{
    public partial class AdjustDedAmt : ComponentBase
    {
        // ── Injected Services ─────────────────────────────────────────────────
        [Inject] private IAdjustDedAmtService    _adjustService       { get; set; } = null!;
        [Inject] private IEmployeeService        _employeeService     { get; set; } = null!;
        [Inject] private INotificationService    _notificationService { get; set; } = null!;
        [Inject] private NavigationManager       _navManager          { get; set; } = null!;

        // ── Query Parameters ──────────────────────────────────────────────────
        [Parameter][SupplyParameterFromQuery] public string? returnUrl   { get; set; }
        [Parameter][SupplyParameterFromQuery] public bool   contractors { get; set; } = false;

        // ── UI State ──────────────────────────────────────────────────────────
        private bool IsLoading          { get; set; } = true;
        private bool IsSaving           { get; set; } = false;
        private bool IsDeductionLoading { get; set; } = false;
        private bool IsContractor       { get; set; } = false;

        // ── Lookups & Selections ──────────────────────────────────────────────
        private List<DepartmentListItemModel> AllDepartments     { get; set; } = new();
        private List<EmployeeListItemModel>   AllEmployees       { get; set; } = new();
        private DepartmentListItemModel?      SelectedDepartment { get; set; }
        private EmployeeListItemModel?        SelectedEmployee   { get; set; }
        private EmpDeductionInfoModel?        DeductionInfo      { get; set; }

        // ── Form Input ────────────────────────────────────────────────────────
        private decimal NewDAmount { get; set; } = 0;

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
            DeductionInfo      = null;
            NewDAmount         = 0;

            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await _employeeService.GetEmployeesAsync(deptId, false, IsContractor);
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            DeductionInfo    = null;
            NewDAmount       = 0;

            if (emp != null)
            {
                IsDeductionLoading = true;
                try
                {
                    DeductionInfo = await _adjustService.GetEmployeeLatestLoanDeductionAsync(emp.EmpID);
                    if (DeductionInfo != null && DeductionInfo.HasActiveLoan)
                    {
                        NewDAmount = DeductionInfo.CurrentDAmount;
                    }
                }
                finally
                {
                    IsDeductionLoading = false;
                }
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
            if (DeductionInfo == null || !DeductionInfo.HasActiveLoan || DeductionInfo.EntryID <= 0)
            {
                _notificationService.ShowWarning("Validation", "No Deduction Feeded For This Employee.");
                return;
            }
            if (NewDAmount < 0)
            {
                _notificationService.ShowWarning("Validation", "Deduction Amount cannot be negative.");
                return;
            }

            var dto = new AdjustDedAmtDto
            {
                EmpID      = SelectedEmployee.EmpID,
                EntryID    = DeductionInfo.EntryID,
                NewDAmount = NewDAmount
            };

            IsSaving = true;
            try
            {
                await _adjustService.UpdateDeductionAmountAsync(dto);

                _notificationService.ShowSuccess(
                    "Saved Successfully",
                    $"Monthly deduction amount updated to {NewDAmount:N2} for [{SelectedEmployee.Name}].");

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
            DeductionInfo    = null;
            NewDAmount       = 0;
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
