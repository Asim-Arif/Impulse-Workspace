using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Impulse.Services.WorkflowTasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.MonthlyAttendance
{
    public partial class MonthlyAttendance : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        // ── Injected Services ──────────────────────────────────────────────
        [Inject] private IMonthlyAttendanceService MonthlyAttendanceService { get; set; } = default!;
        [Inject] private IEmployeeService          EmployeeService          { get; set; } = default!;
        [Inject] private INotificationService      NotificationService      { get; set; } = default!;
        [Inject] private NavigationManager         NavigationManager        { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;
        [Inject] private IAuditService             AuditService             { get; set; } = default!;
        [Inject] private IWorkflowTaskEngine       WorkflowTaskEngine       { get; set; } = default!;
        [Inject] private DataAccessLibrary.Interface.Setup.IUserDataAccess UserDataAccess { get; set; } = default!;
        [Inject] private DataAccessLibrary.Interface.Setup.IUserRoleDataAccess UserRoleDataAccess { get; set; } = default!;

        // ── State ──────────────────────────────────────────────────────────
        private bool IsLoading = false;
        private bool IsSaving  = false;
        private bool IsDirectorOrAdmin = false;
        private bool IsSalaryFinalized = false;
        private string CurrentUserName = "System";

        private DepartmentListItemModel? SelectedDepartment { get; set; }
        private EmployeeListItemModel?   SelectedEmployee   { get; set; }

        private List<DepartmentListItemModel> AllDepartments = new();
        private List<EmployeeListItemModel>   AllEmployees   = new();
        private List<MonthlyAttendanceDayRow> DayRows        = new();
        private List<MonthlyAttendanceDayRow> OriginalDayRows = new();

        private int CurrentYear  = DateTime.Today.Year;
        private int CurrentMonth = DateTime.Today.Month;
        private string SelectedMonthYear => $"{CurrentYear:D4}-{CurrentMonth:D2}";

        // Reason Modal State for Workflow
        private bool ShowReasonModal = false;
        private string ReasonModalTitle = "Monthly Attendance Modification Request";
        private string ReasonText = string.Empty;
        private bool IsSubmittingRequest = false;
        private AttendanceWorkflowActionType PendingActionType;
        private MonthlyAttendanceDayRow? PendingClearRow = null;

        // ── Initialization ─────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                    var user = authState.User;
                    
                    bool isRoleMatch = user.IsInRole("Director") || user.IsInRole("Admin") || user.IsInRole("Administrator") || user.IsInRole("SuperAdmin");
                    bool isNameMatch = string.Equals(CurrentUserName, "admin", StringComparison.OrdinalIgnoreCase) || string.Equals(CurrentUserName, "administrator", StringComparison.OrdinalIgnoreCase);

                    if (isRoleMatch || isNameMatch)
                    {
                        IsDirectorOrAdmin = true;
                    }
                    else
                    {
                        var dbUser = await UserDataAccess.GetUserByUserNameAsync(CurrentUserName);
                        if (dbUser != null)
                        {
                            if (dbUser.UserManagement == true)
                            {
                                IsDirectorOrAdmin = true;
                            }
                            else
                            {
                                var userRoles = await UserRoleDataAccess.GetRolesByUserIdAsync(dbUser.UserID);
                                if (userRoles != null && userRoles.Any(r => r.Equals("Director", StringComparison.OrdinalIgnoreCase) || r.Equals("Admin", StringComparison.OrdinalIgnoreCase) || r.Equals("Administrator", StringComparison.OrdinalIgnoreCase)))
                                {
                                    IsDirectorOrAdmin = true;
                                }
                            }
                        }
                    }
                }

                AllDepartments = await EmployeeService.GetDepartmentsAsync(false);
                AllEmployees   = await EmployeeService.GetEmployeesAsync("0", false, false);
                await CheckSalaryLockAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        private async Task CheckSalaryLockAsync()
        {
            IsSalaryFinalized = await MonthlyAttendanceService.IsSalaryFinalizedAsync(CurrentYear, CurrentMonth);
        }

        // ── Filters & Search ───────────────────────────────────────────────
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

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await EmployeeService.GetEmployeesAsync(deptId, false, false);
            SelectedEmployee = null;
            DayRows = new();
            OriginalDayRows = new();
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            await LoadMonthGridAsync();
        }

        private async Task OnMonthYearChanged(ChangeEventArgs e)
        {
            string? val = e?.Value?.ToString();
            if (!string.IsNullOrWhiteSpace(val) && DateTime.TryParse($"{val}-01", out DateTime dt))
            {
                CurrentYear  = dt.Year;
                CurrentMonth = dt.Month;
                await CheckSalaryLockAsync();
                await LoadMonthGridAsync();
            }
        }

        private async Task LoadMonthGridAsync()
        {
            if (SelectedEmployee == null || string.IsNullOrWhiteSpace(SelectedEmployee.EmpID) || SelectedEmployee.EmpID == "0")
            {
                DayRows = new();
                OriginalDayRows = new();
                return;
            }

            IsLoading = true;
            StateHasChanged();

            try
            {
                await CheckSalaryLockAsync();
                DayRows = await MonthlyAttendanceService.GetMonthlyAttendanceAsync(SelectedEmployee.EmpID, CurrentYear, CurrentMonth);
                
                // Keep deep copy of original rows for Before vs After comparison
                OriginalDayRows = DayRows.Select(r => new MonthlyAttendanceDayRow
                {
                    DayNo = r.DayNo,
                    Date = r.Date,
                    DayName = r.DayName,
                    IsSundayOrHoliday = r.IsSundayOrHoliday,
                    HolidayName = r.HolidayName,
                    Status = r.Status,
                    InTime = r.InTime,
                    OutTime = r.OutTime,
                    OtHours = r.OtHours,
                    IsModified = false
                }).ToList();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Attendance", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ── Attendance Editing ─────────────────────────────────────────────
        private void OnInTimeChanged(MonthlyAttendanceDayRow row, ChangeEventArgs e)
        {
            row.InTime = e?.Value?.ToString() ?? string.Empty;
            row.IsModified = true;
            if (!string.IsNullOrWhiteSpace(row.InTime) && row.Status == "Absent")
            {
                row.Status = "Present";
            }
        }

        private void OnOutTimeChanged(MonthlyAttendanceDayRow row, ChangeEventArgs e)
        {
            row.OutTime = e?.Value?.ToString() ?? string.Empty;
            row.IsModified = true;
        }

        private async Task ClearDateAsync(MonthlyAttendanceDayRow row)
        {
            if (SelectedEmployee == null) return;

            if (IsSalaryFinalized)
            {
                NotificationService.ShowError("Locked", $"Salary for {SelectedMonthYear} has already been finalized in MonthlySalaries. Attendance cannot be cleared.");
                return;
            }

            // If user is Director/Admin -> Direct Bypass
            if (IsDirectorOrAdmin)
            {
                await ExecuteDirectClearDateAsync(row);
                return;
            }

            // Regular operator -> Request Approval
            PendingClearRow = row;
            PendingActionType = AttendanceWorkflowActionType.MonthlyClearDate;
            ReasonModalTitle = $"Request Clear Attendance - [{SelectedEmployee.EmpID}] {SelectedEmployee.Name} ({row.Date:dd-MMM-yyyy})";
            ReasonText = string.Empty;
            ShowReasonModal = true;
        }

        private async Task ExecuteDirectClearDateAsync(MonthlyAttendanceDayRow row)
        {
            if (SelectedEmployee == null) return;
            try
            {
                bool success = await MonthlyAttendanceService.ClearDateAttendanceAsync(SelectedEmployee.EmpID, row.Date);
                if (success)
                {
                    row.InTime = string.Empty;
                    row.OutTime = string.Empty;
                    row.Status = "Absent";
                    row.OtHours = 0;
                    row.IsModified = false;
                    NotificationService.ShowSuccess("Cleared", $"Attendance cleared for {row.Date:dd-MMM-yyyy}.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Clear Failed", ex.Message);
            }
        }

        private async Task SaveMonthAsync()
        {
            if (SelectedEmployee == null || string.IsNullOrWhiteSpace(SelectedEmployee.EmpID)) return;

            if (IsSalaryFinalized)
            {
                NotificationService.ShowError("Locked", $"Salary for {SelectedMonthYear} has already been finalized in MonthlySalaries. Attendance cannot be saved.");
                return;
            }

            var modifiedRows = DayRows.Where(r => r.IsModified).ToList();
            if (!modifiedRows.Any())
            {
                NotificationService.ShowWarning("No Changes", "No attendance records have been modified.");
                return;
            }

            // If user is Director/Admin -> Direct Bypass
            if (IsDirectorOrAdmin)
            {
                await ExecuteDirectSaveMonthAsync();
                return;
            }

            // Regular operator -> Request 1 Consolidated Monthly Approval
            PendingClearRow = null;
            PendingActionType = AttendanceWorkflowActionType.MonthlySave;
            ReasonModalTitle = $"Request Monthly Attendance Save - [{SelectedEmployee.EmpID}] {SelectedEmployee.Name} ({SelectedMonthYear})";
            ReasonText = string.Empty;
            ShowReasonModal = true;
        }

        private async Task ExecuteDirectSaveMonthAsync()
        {
            if (SelectedEmployee == null || string.IsNullOrWhiteSpace(SelectedEmployee.EmpID)) return;

            IsSaving = true;
            StateHasChanged();

            try
            {
                var input = new MonthlyAttendanceSaveDto
                {
                    EmpID = SelectedEmployee.EmpID,
                    Year  = CurrentYear,
                    Month = CurrentMonth,
                    Rows  = DayRows
                };

                bool success = await MonthlyAttendanceService.SaveMonthlyAttendanceAsync(input);
                if (success)
                {
                    NotificationService.ShowSuccess("Saved", "Monthly attendance saved successfully.");
                    await LoadMonthGridAsync();
                }
                else
                {
                    NotificationService.ShowError("Save Failed", "Could not save monthly attendance.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        private async Task SubmitWorkflowRequestAsync()
        {
            if (SelectedEmployee == null) return;
            if (string.IsNullOrWhiteSpace(ReasonText))
            {
                NotificationService.ShowWarning("Reason Required", "Please enter a reason for this attendance modification request.");
                return;
            }

            IsSubmittingRequest = true;
            StateHasChanged();

            try
            {
                string clientIp = AuditService.GetClientIpAddress();
                var diffList = new List<AttendanceDiffRow>();

                if (PendingActionType == AttendanceWorkflowActionType.MonthlyClearDate && PendingClearRow != null)
                {
                    var orig = OriginalDayRows.FirstOrDefault(r => r.Date.Date == PendingClearRow.Date.Date);
                    diffList.Add(new AttendanceDiffRow
                    {
                        Date = PendingClearRow.Date,
                        DayName = PendingClearRow.DayName,
                        PrevInTime = orig?.InTime ?? string.Empty,
                        PrevOutTime = orig?.OutTime ?? string.Empty,
                        PrevStatus = orig?.Status ?? "Absent",
                        PrevOtHours = orig?.OtHours ?? 0,
                        NewInTime = string.Empty,
                        NewOutTime = string.Empty,
                        NewStatus = "Absent",
                        NewOtHours = 0,
                        IsDeleted = true
                    });
                }
                else
                {
                    // MonthlySave: Include all modified rows in diff
                    var modifiedRows = DayRows.Where(r => r.IsModified).ToList();
                    foreach (var modRow in modifiedRows)
                    {
                        var orig = OriginalDayRows.FirstOrDefault(r => r.Date.Date == modRow.Date.Date);
                        diffList.Add(new AttendanceDiffRow
                        {
                            Date = modRow.Date,
                            DayName = modRow.DayName,
                            PrevInTime = orig?.InTime ?? string.Empty,
                            PrevOutTime = orig?.OutTime ?? string.Empty,
                            PrevStatus = orig?.Status ?? "Absent",
                            PrevOtHours = orig?.OtHours ?? 0,
                            NewInTime = modRow.InTime,
                            NewOutTime = modRow.OutTime,
                            NewStatus = modRow.Status,
                            NewOtHours = modRow.OtHours,
                            IsDeleted = false
                        });
                    }
                }

                var monthlyInput = new MonthlyAttendanceSaveDto
                {
                    EmpID = SelectedEmployee.EmpID,
                    Year = CurrentYear,
                    Month = CurrentMonth,
                    Rows = DayRows
                };

                var request = new AttendanceWorkflowRequestDto
                {
                    ActionType = PendingActionType,
                    EmpID = SelectedEmployee.EmpID,
                    EmployeeName = SelectedEmployee.Name,
                    DepartmentName = SelectedEmployee.Designation ?? string.Empty,
                    Year = CurrentYear,
                    Month = CurrentMonth,
                    AttendanceDate = PendingActionType == AttendanceWorkflowActionType.MonthlyClearDate ? PendingClearRow?.Date : null,
                    Reason = ReasonText.Trim(),
                    OriginatorUserName = CurrentUserName,
                    MachineName = clientIp,
                    MonthlyInput = monthlyInput,
                    DiffRows = diffList
                };

                int taskId = await WorkflowTaskEngine.RequestAttendanceActionAsync(request);

                NotificationService.ShowSuccess("Request Submitted", $"Monthly attendance change request has been submitted to the Director for approval (Task #{taskId}).");
                ShowReasonModal = false;
                ReasonText = string.Empty;
                await LoadMonthGridAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Submission Failed", ex.Message);
            }
            finally
            {
                IsSubmittingRequest = false;
                StateHasChanged();
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
                NavigationManager.NavigateTo(ReturnUrl);
            else
                NavigationManager.NavigateTo("/payroll");
        }
    }
}
