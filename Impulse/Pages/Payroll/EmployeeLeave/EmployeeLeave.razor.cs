using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.EmployeeLeave
{
    public partial class EmployeeLeave : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        // ── Injected Services ──────────────────────────────────────────────
        [Inject] private IEmployeeLeaveService   EmployeeLeaveService     { get; set; } = default!;
        [Inject] private IEmployeeService        EmployeeService          { get; set; } = default!;
        [Inject] private INotificationService    NotificationService      { get; set; } = default!;
        [Inject] private NavigationManager       NavigationManager        { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;

        // ── State ──────────────────────────────────────────────────────────
        private bool IsLoading = false;
        private bool IsSaving  = false;

        private DepartmentListItemModel? SelectedDepartment { get; set; }
        private EmployeeListItemModel?   SelectedEmployee   { get; set; }
        private UserLookupModel?         SelectedAuthUser   { get; set; }

        private List<DepartmentListItemModel> AllDepartments = new();
        private List<EmployeeListItemModel>   AllEmployees   = new();
        private List<UserLookupModel>         AllUsers       = new();
        private List<LeaveTypeLookupModel>    LeaveTypes     = new();

        private LeaveBalanceMatrixDto Balances = new();
        private EmployeeLeaveInputDto Input    = new();

        // ── Initialization ─────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            try
            {
                AllDepartments = await EmployeeService.GetDepartmentsAsync(false);
                AllEmployees   = await EmployeeService.GetEmployeesAsync("0", false, false);
                AllUsers       = await EmployeeLeaveService.GetUsersAsync();

                InitLeaveTypes();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        private void InitLeaveTypes()
        {
            LeaveTypes = new List<LeaveTypeLookupModel>
            {
                new() { Code = "CL",  Name = "Casual Leave",           Limit = 10 },
                new() { Code = "SL",  Name = "Sick Leave",             Limit = 8  },
                new() { Code = "AL",  Name = "Annual Leave",           Limit = 14 },
                new() { Code = "CPL", Name = "Compensatory Leave",     Limit = 10 },
                new() { Code = "LWP", Name = "Unapproved Leave (LWP)", Limit = 0  },
                new() { Code = "ML",  Name = "Maternity Leave",        Limit = 90 },
                new() { Code = "SAL", Name = "Special Approved Leave", Limit = 0  }
            };
        }

        // ── Filter Handlers ────────────────────────────────────────────────
        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AllDepartments;
            return await Task.FromResult(
                AllDepartments.Where(d => d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AllEmployees;

            return await Task.FromResult(
                AllEmployees.Where(e =>
                    e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrEmpty(e.FName) && e.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(e.Designation) && e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                )
            );
        }

        private async Task<IEnumerable<UserLookupModel>> SearchUsers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AllUsers;
            return await Task.FromResult(
                AllUsers.Where(u => u.UserName.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await EmployeeService.GetEmployeesAsync(deptId, false, false);
            SelectedEmployee = null;
            Balances = new LeaveBalanceMatrixDto();
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            if (emp != null && emp.EmpID != "0")
            {
                Input.EmpID = emp.EmpID;
                await LoadBalancesAsync();
            }
            else
            {
                Input.EmpID = string.Empty;
                Balances = new LeaveBalanceMatrixDto();
            }
        }

        private async Task LoadBalancesAsync()
        {
            if (string.IsNullOrWhiteSpace(Input.EmpID)) return;
            try
            {
                Balances = await EmployeeLeaveService.GetLeaveBalancesAsync(Input.EmpID, Input.FromDate);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Balances", ex.Message);
            }
        }

        private void OnAuthUserChanged(UserLookupModel? user)
        {
            SelectedAuthUser = user;
            Input.AuthPerson = user?.UserName ?? string.Empty;
        }

        private void SetEntryMode(bool isAbsent)
        {
            Input.IsAbsent = isAbsent;
        }

        private void SetDurationMode(string mode)
        {
            Input.DurationMode = mode;
        }

        private async Task OnFromDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime dt))
            {
                Input.FromDate = dt;
                if (Input.ToDate < Input.FromDate)
                    Input.ToDate = Input.FromDate;

                await LoadBalancesAsync();
            }
        }

        private void OnToDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime dt))
            {
                Input.ToDate = dt;
            }
        }

        private void OnLeaveTypeChanged(ChangeEventArgs e)
        {
            string code = e?.Value?.ToString() ?? "CL";
            Input.LeaveTypeCode = code;
            var lt = LeaveTypes.FirstOrDefault(l => l.Code == code);
            Input.LeaveTypeName = lt?.Name ?? code;
        }

        // ── Form Validation & Saving ───────────────────────────────────────
        private async Task<bool> PerformSaveAsync()
        {
            if (SelectedEmployee == null || string.IsNullOrWhiteSpace(Input.EmpID))
            {
                NotificationService.ShowWarning("Validation Warning", "Please select a valid employee.");
                return false;
            }

            if (Input.ToDate < Input.FromDate)
            {
                NotificationService.ShowWarning("Validation Warning", "Invalid date range. 'To Date' cannot be earlier than 'From Date'.");
                return false;
            }

            if (!Input.IsAbsent)
            {
                if (SelectedAuthUser == null || string.IsNullOrWhiteSpace(Input.AuthPerson))
                {
                    NotificationService.ShowWarning("Validation Warning", "Please select an Authority Person.");
                    return false;
                }

                if (string.Equals(SelectedAuthUser.UserName, Input.EmpID, StringComparison.OrdinalIgnoreCase))
                {
                    NotificationService.ShowWarning("Validation Warning", "Employee and Authority Person cannot be the same.");
                    return false;
                }

                // Probation check
                bool inProbation = await EmployeeLeaveService.IsProbationAsync(Input.EmpID, Input.FromDate);
                if (inProbation)
                {
                    NotificationService.ShowWarning("Probation Restriction", "Employee is still in Probation. Paid leave cannot be applied.");
                    return false;
                }
            }

            IsSaving = true;
            StateHasChanged();

            try
            {
                bool success = await EmployeeLeaveService.SaveEmployeeLeaveAsync(Input);
                if (success)
                {
                    NotificationService.ShowSuccess("Success", Input.IsAbsent ? "Absence recorded successfully." : "Leave recorded successfully.");
                    await LoadBalancesAsync();
                    return true;
                }
                else
                {
                    NotificationService.ShowError("Save Failed", "Could not save leave entry.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Error", ex.Message);
                return false;
            }
            finally
            {
                IsSaving = false;
            }
        }

        private async Task SaveAsync()
        {
            await PerformSaveAsync();
        }

        private async Task SaveAndNextAsync()
        {
            bool saved = await PerformSaveAsync();
            if (saved)
            {
                ResetForm();
            }
        }

        private void ResetForm()
        {
            SelectedEmployee = null;
            SelectedAuthUser = null;
            Input = new EmployeeLeaveInputDto();
            Balances = new LeaveBalanceMatrixDto();
        }

        // ── Reporting ─────────────────────────────────────────────────────
        private async Task PrintLeaveReportAsync()
        {
            try
            {
                await EmployeeLeaveService.ExecutePrintEmpLeaveBalancesAsync(Input.FromDate);

                string selFormula = SelectedEmployee != null && SelectedEmployee.EmpID != "0"
                    ? $"{{VEmpLeaveBalance.EmpID}}='{SelectedEmployee.EmpID}'"
                    : string.Empty;

                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.Payroll.EmpLeaveBalanceMonthly,
                    SelectionFormula = selFormula,
                    Parameters = new Dictionary<string, object>
                    {
                        { "@ForDate", Input.FromDate.ToString("dd-MMM-yyyy") }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        // ── Navigation ────────────────────────────────────────────────────
        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
                NavigationManager.NavigateTo(ReturnUrl);
            else
                NavigationManager.NavigateTo("/payroll");
        }
    }
}
