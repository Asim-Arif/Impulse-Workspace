using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.AdvanceLedger
{
    public partial class AdvanceLedger : ComponentBase
    {
        // ── Injected Services ─────────────────────────────────────────────────
        [Inject] private IAdvanceLedgerService   _ledgerService       { get; set; } = null!;
        [Inject] private IEmployeeService        _employeeService     { get; set; } = null!;
        [Inject] private INotificationService    _notificationService { get; set; } = null!;
        [Inject] private NavigationManager       _navManager          { get; set; } = null!;

        // ── Query Parameters ──────────────────────────────────────────────────
        [Parameter][SupplyParameterFromQuery] public string? returnUrl   { get; set; }
        [Parameter][SupplyParameterFromQuery] public bool   contractors { get; set; } = false;

        // ── UI State ──────────────────────────────────────────────────────────
        private bool IsLoading    { get; set; } = true;
        private bool IsContractor { get; set; } = false;

        // ── Filter State ──────────────────────────────────────────────────────
        private AdvanceLedgerFilterDto Filter { get; set; } = new();
        private string StatusFilter           { get; set; } = "All"; // "All", "Active", "Cleared"
        private string SearchQuery            { get; set; } = string.Empty;

        // ── Lookups & Selections ──────────────────────────────────────────────
        private List<DepartmentListItemModel> AllDepartments     { get; set; } = new();
        private List<EmployeeListItemModel>   AllEmployees       { get; set; } = new();
        private DepartmentListItemModel?      SelectedDepartment { get; set; }
        private EmployeeListItemModel?        SelectedEmployee   { get; set; }

        // ── Data Items & Metrics ──────────────────────────────────────────────
        private List<AdvanceLedgerItemModel> LedgerItems { get; set; } = new();

        private List<AdvanceLedgerItemModel> FilteredLedgerItems
        {
            get
            {
                var query = LedgerItems.AsEnumerable();

                // Status Filter
                if (StatusFilter == "Active")
                {
                    query = query.Where(i => i.Cleared == 0 && i.Balance > 0);
                }
                else if (StatusFilter == "Cleared")
                {
                    query = query.Where(i => i.Cleared != 0 || i.Balance <= 0);
                }

                // Quick Search Text
                if (!string.IsNullOrWhiteSpace(SearchQuery))
                {
                    string q = SearchQuery.Trim();
                    query = query.Where(i =>
                        i.EmpID.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        i.EmpName.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        (i.AccVoucherNo ?? "").Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        (i.DeptName ?? "").Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        (i.Designation ?? "").Contains(q, StringComparison.OrdinalIgnoreCase)
                    );
                }

                return query.ToList();
            }
        }

        private AdvanceLedgerKpiSummaryModel KpiSummary
        {
            get
            {
                var items = FilteredLedgerItems;
                return new AdvanceLedgerKpiSummaryModel
                {
                    TotalIssued      = items.Sum(i => i.Amount),
                    TotalCleared     = items.Sum(i => i.AmountCleared),
                    TotalOutstanding = items.Sum(i => i.Balance)
                };
            }
        }

        // ── Lifecycle ─────────────────────────────────────────────────────────

        protected override async Task OnInitializedAsync()
        {
            IsContractor = contractors;
            Filter.IsContractor = IsContractor;

            try
            {
                AllDepartments = await _employeeService.GetDepartmentsAsync(false);
                AllEmployees   = await _employeeService.GetEmployeesAsync("0", false, IsContractor);

                await LoadLedgerAsync();
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

        // ── Data Fetch ────────────────────────────────────────────────────────

        private async Task LoadLedgerAsync()
        {
            IsLoading = true;
            try
            {
                LedgerItems = await _ledgerService.GetAdvanceLedgerAsync(Filter);
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Fetch Error", ex.Message);
                LedgerItems = new List<AdvanceLedgerItemModel>();
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task RefreshAsync()
        {
            await LoadLedgerAsync();
        }

        // ── Filter Handlers ───────────────────────────────────────────────────

        private async Task SetLoanType(int loanType)
        {
            Filter.LoanType = loanType;
            await LoadLedgerAsync();
        }

        private async Task OnFromDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime d))
            {
                Filter.FromDate = d;
                await LoadLedgerAsync();
            }
        }

        private async Task OnToDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime d))
            {
                Filter.ToDate = d;
                await LoadLedgerAsync();
            }
        }

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            SelectedEmployee   = null;
            Filter.DeptID      = dept?.DeptID ?? "0";
            Filter.EmpID       = "0";

            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await _employeeService.GetEmployeesAsync(deptId, false, IsContractor);

            await LoadLedgerAsync();
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            Filter.EmpID     = emp?.EmpID ?? "0";
            await LoadLedgerAsync();
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
