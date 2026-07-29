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
using BlazorContextMenu;

namespace Impulse.Pages.Payroll.BiometricLedger
{
    public partial class BiometricLedger : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        // ── Injected Services ──────────────────────────────────────────────
        [Inject] private IBiometricLedgerService  BiometricLedgerService  { get; set; } = default!;
        [Inject] private IIFaceMachineService     IFaceMachineService     { get; set; } = default!;
        [Inject] private IEmployeeService         EmployeeService         { get; set; } = default!;
        [Inject] private INotificationService     NotificationService     { get; set; } = default!;
        [Inject] private NavigationManager        NavigationManager       { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;

        // ── State ──────────────────────────────────────────────────────────
        private bool IsLoading       = false;
        private bool IsReadingMachine = false;

        private DateTime SelectedDate { get; set; } = DateTime.Today;

        private DepartmentListItemModel?  SelectedDepartment { get; set; }
        private EmployeeListItemModel?    SelectedEmployee   { get; set; }

        private List<DepartmentListItemModel>  AllDepartments = new();
        private List<EmployeeListItemModel>    AllEmployees   = new();
        private List<BiometricMachineItem>     IFaceMachines  = new();

        private List<BiometricDeptSummaryRow>  DeptSummaries  = new();
        private List<BiometricEmployeeRow>     EmployeeRows   = new();

        // Tracks which dept IDs are currently expanded (starts with all expanded)
        private HashSet<string> ExpandedDepts = new();

        // IFace machine panel
        private int SelectedMachineId { get; set; } = 0;

        // ── Initialization ─────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            try
            {
                AllDepartments = await EmployeeService.GetDepartmentsAsync(false);
                AllEmployees   = await EmployeeService.GetEmployeesAsync("0", false, false);
                IFaceMachines  = await BiometricLedgerService.GetIFaceMachinesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        // ── Data Loading ───────────────────────────────────────────────────
        private async Task LoadDataAsync()
        {
            IsLoading = true;
            StateHasChanged();
            try
            {
                string? deptId = SelectedDepartment?.DeptID;
                string? empId  = SelectedEmployee?.EmpID;

                DeptSummaries = await BiometricLedgerService.GetDeptSummaryAsync(SelectedDate, deptId);
                EmployeeRows  = await BiometricLedgerService.GetEmployeeAttendanceAsync(SelectedDate, deptId, empId);

                // All dept rows start expanded
                ExpandedDepts = new HashSet<string>(DeptSummaries.Select(d => d.DeptID));
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Load Error", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
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
                    (e.FName != null && e.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                    (e.Designation != null && e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                )
            );
        }

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            await LoadDataAsync();
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            await LoadDataAsync();
        }

        private async Task OnDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime parsed))
            {
                SelectedDate = parsed;
                await LoadDataAsync();
            }
        }

        // ── Expand / Collapse ─────────────────────────────────────────────
        private void ToggleDept(string deptId)
        {
            if (ExpandedDepts.Contains(deptId))
                ExpandedDepts.Remove(deptId);
            else
                ExpandedDepts.Add(deptId);
        }

        // ── IFace Machine Reading ─────────────────────────────────────────
        private async Task ReadFromIFaceMachineAsync()
        {
            if (SelectedMachineId == 0) return;

            var machine = IFaceMachines.FirstOrDefault(m => m.EntryID == SelectedMachineId);
            if (machine == null) return;

            IsReadingMachine = true;
            StateHasChanged();
            try
            {
                var (recordsRead, recordsProcessed, message) =
                    await IFaceMachineService.ReadAndProcessAsync(machine.IPAddress, machine.MachineNo);

                if (recordsProcessed > 0)
                    NotificationService.ShowSuccess("Machine Read Complete", message);
                else
                    NotificationService.ShowWarning("Machine Read", message);

                // Refresh the ledger to reflect newly imported punches
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("IFace Machine Error", ex.Message);
            }
            finally
            {
                IsReadingMachine = false;
            }
        }

        // ── Options Menu Trigger ──────────────────────────────────────────
        private void OpenOptionsMenu(MouseEventArgs e)
        {
            // BlazorContextMenu handles placement from the trigger; no manual override needed
        }

        // ── Reports ───────────────────────────────────────────────────────
        private async Task PrintDailyAttendance(ItemClickEventArgs e)
        {
            try
            {
                await BiometricLedgerService.ExecutePrintDailyAbsentLeaveDetailAsync(SelectedDate);

                string selFormula = "{Employees.Active}=True AND {VMainDepartments.TempDept}=False";
                if (SelectedDepartment != null)
                    selFormula += $" AND {{VMainDepartments.DeptID}}='{SelectedDepartment.DeptID}'";

                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.Payroll.DailyAttendance,
                    SelectionFormula = selFormula,
                    Parameters = new Dictionary<string, object>
                    {
                        { "@ForDate", SelectedDate.ToString("dd-MMM-yyyy") }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private async Task PrintAbsentLeaveSummary(ItemClickEventArgs e)
        {
            try
            {
                await BiometricLedgerService.ExecutePrintDailyAbsentLeaveSummaryAsync(SelectedDate);
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.Payroll.AbsentLeaveSummary,
                    Parameters = new Dictionary<string, object>
                    {
                        { "@ForDate", SelectedDate.ToString("dd-MMM-yyyy") }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private async Task PrintLateComers(ItemClickEventArgs e)
        {
            try
            {
                await BiometricLedgerService.ExecutePrintDailyAbsentLeaveDetailAsync(SelectedDate);
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.Payroll.LateComers,
                    SelectionFormula = "{PrintDailyAbsentLeaveDetail.LateComing}=TRUE",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@ForDate", SelectedDate.ToString("dd-MMM-yyyy") }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private async Task PrintPresentEmployees(ItemClickEventArgs e)
        {
            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.Payroll.PresentEmpList,
                    SelectionFormula = $"{{AttendanceSheet.DT}}=#{SelectedDate:dd-MMM-yyyy}# AND {{AttendanceSheet.Attendance}}=1",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@ForDate", SelectedDate.ToString("dd-MMM-yyyy") }
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

