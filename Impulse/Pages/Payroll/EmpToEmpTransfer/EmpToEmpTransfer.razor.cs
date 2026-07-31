using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.EmpToEmpTransfer
{
    public partial class EmpToEmpTransfer : ComponentBase
    {
        // ── Injected Services ─────────────────────────────────────────────────
        [Inject] private IEmpToEmpTransferService _transferService     { get; set; } = null!;
        [Inject] private IEmployeeService         _employeeService     { get; set; } = null!;
        [Inject] private INotificationService     _notificationService { get; set; } = null!;
        [Inject] private NavigationManager        _navManager          { get; set; } = null!;

        // ── Query Parameters ──────────────────────────────────────────────────
        [Parameter][SupplyParameterFromQuery] public string? returnUrl   { get; set; }
        [Parameter][SupplyParameterFromQuery] public bool   contractors { get; set; } = false;

        // ── UI State ──────────────────────────────────────────────────────────
        private bool IsLoading    { get; set; } = true;
        private bool IsSaving     { get; set; } = false;
        private bool IsContractor { get; set; } = false;

        // ── Source Employee (From) ────────────────────────────────────────────
        private List<DepartmentListItemModel> FromDepartments        { get; set; } = new();
        private List<EmployeeListItemModel>   FromEmployees          { get; set; } = new();
        private DepartmentListItemModel?      SelectedFromDepartment { get; set; }
        private EmployeeListItemModel?        SelectedFromEmployee   { get; set; }
        private EmpTransferBalanceSummaryModel? FromEmployeeSummary  { get; set; }

        // ── Destination Employee (To) ─────────────────────────────────────────
        private List<DepartmentListItemModel> ToDepartments        { get; set; } = new();
        private List<EmployeeListItemModel>   ToEmployees          { get; set; } = new();
        private DepartmentListItemModel?      SelectedToDepartment { get; set; }
        private EmployeeListItemModel?        SelectedToEmployee   { get; set; }

        // ── Transfer Parameters ───────────────────────────────────────────────
        private DateTime TransferDate     { get; set; } = DateTime.Today;
        private DateTime DeductionStartDT { get; set; } = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        private decimal  LTAmt            { get; set; } = 0;
        private decimal  LTDeduct         { get; set; } = 0;
        private decimal  STAmt            { get; set; } = 0;
        private string   Description      { get; set; } = "Employee to Employee Loan Transfer";

        // ── Lifecycle ─────────────────────────────────────────────────────────

        protected override async Task OnInitializedAsync()
        {
            IsContractor = contractors;
            IsLoading    = true;

            try
            {
                // Load department lists
                FromDepartments = await _employeeService.GetDepartmentsAsync(false);
                ToDepartments   = new List<DepartmentListItemModel>(FromDepartments);

                // Load employee lists (all depts initially)
                FromEmployees = await _employeeService.GetEmployeesAsync("0", false, IsContractor);
                ToEmployees   = new List<EmployeeListItemModel>(FromEmployees);
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

        // ── Source Employee Event Handlers ────────────────────────────────────

        private async Task OnFromDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedFromDepartment = dept;
            SelectedFromEmployee   = null;
            FromEmployeeSummary    = null;

            string deptId = dept?.DeptID ?? "0";
            FromEmployees = await _employeeService.GetEmployeesAsync(deptId, false, IsContractor);
        }

        private async Task OnFromEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedFromEmployee = emp;
            if (emp != null)
            {
                FromEmployeeSummary = await _transferService.GetEmployeeBalancesAsync(emp.EmpID);
            }
            else
            {
                FromEmployeeSummary = null;
            }
        }

        // ── Destination Employee Event Handlers ──────────────────────────────

        private async Task OnToDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedToDepartment = dept;
            SelectedToEmployee   = null;

            string deptId = dept?.DeptID ?? "0";
            ToEmployees = await _employeeService.GetEmployeesAsync(deptId, false, IsContractor);
        }

        private void OnToEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedToEmployee = emp;
        }

        // ── Date Event Handlers ───────────────────────────────────────────────

        private void OnTransferDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime d))
            {
                TransferDate     = d;
                DeductionStartDT = new DateTime(d.Year, d.Month, 1);
            }
        }

        private void OnDeductionStartChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString() + "-01", out DateTime d))
            {
                DeductionStartDT = new DateTime(d.Year, d.Month, 1);
            }
        }

        // ── Typeahead Search Methods ──────────────────────────────────────────

        private Task<IEnumerable<DepartmentListItemModel>> SearchFromDepartments(string search)
        {
            var results = FromDepartments
                .Where(d => string.IsNullOrWhiteSpace(search) ||
                            d.Name.Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        private Task<IEnumerable<EmployeeListItemModel>> SearchFromEmployees(string search)
        {
            var results = FromEmployees
                .Where(e => string.IsNullOrWhiteSpace(search) ||
                            e.EmpID.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            e.Name.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            (e.FName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            (e.Designation ?? "").Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        private Task<IEnumerable<DepartmentListItemModel>> SearchToDepartments(string search)
        {
            var results = ToDepartments
                .Where(d => string.IsNullOrWhiteSpace(search) ||
                            d.Name.Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        private Task<IEnumerable<EmployeeListItemModel>> SearchToEmployees(string search)
        {
            var results = ToEmployees
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

            // ── Validation ────────────────────────────────────────────────────
            if (SelectedFromEmployee == null)
            {
                _notificationService.ShowWarning("Validation", "Please select Source Employee (Transfer From).");
                return;
            }
            if (SelectedToEmployee == null)
            {
                _notificationService.ShowWarning("Validation", "Please select Destination Employee (Transfer To).");
                return;
            }
            if (SelectedFromEmployee.EmpID == SelectedToEmployee.EmpID)
            {
                _notificationService.ShowWarning("Validation", "Source and Destination Employee cannot be the same person.");
                return;
            }
            if (LTAmt <= 0 && STAmt <= 0)
            {
                _notificationService.ShowWarning("Validation", "Please enter L.T. Transfer Amount or S.T. Transfer Amount.");
                return;
            }
            if (FromEmployeeSummary != null)
            {
                if (LTAmt > FromEmployeeSummary.LTBalance)
                {
                    _notificationService.ShowWarning("Validation", $"L.T. Transfer Amount ({LTAmt:N2}) exceeds Source Employee's balance ({FromEmployeeSummary.LTBalance:N2}).");
                    return;
                }
                if (STAmt > FromEmployeeSummary.STBalance)
                {
                    _notificationService.ShowWarning("Validation", $"S.T. Transfer Amount ({STAmt:N2}) exceeds Source Employee's balance ({FromEmployeeSummary.STBalance:N2}).");
                    return;
                }
            }

            // ── Build DTO ─────────────────────────────────────────────────────
            var dto = new EmpToEmpTransferDto
            {
                FromEmpID        = SelectedFromEmployee.EmpID,
                ToEmpID          = SelectedToEmployee.EmpID,
                TransferDate     = TransferDate,
                DeductionStartDT = DeductionStartDT,
                LTAmt            = LTAmt,
                LTDeduct         = LTDeduct,
                STAmt            = STAmt,
                Description      = string.IsNullOrWhiteSpace(Description) ? "Employee to Employee Loan Transfer" : Description,
                IsContractor     = IsContractor
            };

            // ── Save ──────────────────────────────────────────────────────────
            IsSaving = true;
            try
            {
                await _transferService.ExecuteTransferAsync(dto);

                _notificationService.ShowSuccess(
                    "Transfer Completed",
                    $"Loan/Advance transferred from [{SelectedFromEmployee.Name}] to [{SelectedToEmployee.Name}].");

                // Reset form
                await ResetForm();
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Transfer Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        // ── Reset Form ────────────────────────────────────────────────────────

        private async Task ResetForm()
        {
            SelectedFromEmployee   = null;
            SelectedToEmployee     = null;
            FromEmployeeSummary    = null;
            LTAmt                  = 0;
            LTDeduct               = 0;
            STAmt                  = 0;
            Description            = "Employee to Employee Loan Transfer";
            TransferDate           = DateTime.Today;
            DeductionStartDT       = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
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
