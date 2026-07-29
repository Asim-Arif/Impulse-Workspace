using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Components.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.ManualAttendance
{
    public partial class ManualAttendance : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        private bool IsSaving = false;
        private bool IsLoadingDetails = false;
        private string CurrentUserName = "System";

        private DateTime AttendanceDate { get; set; } = DateTime.Today;
        private bool UseServerTime { get; set; } = false;

        private TimeOnly? InTime { get; set; } = new TimeOnly(8, 0);
        private TimeOnly? BrkOut { get; set; } = new TimeOnly(13, 0);
        private TimeOnly? BrkIn { get; set; } = new TimeOnly(14, 0);
        private TimeOnly? ChkOut { get; set; } = new TimeOnly(17, 0);

        private bool EnableOT { get; set; } = false;
        private TimeOnly? OTIn { get; set; } = new TimeOnly(17, 30);
        private TimeOnly? OTOut { get; set; } = new TimeOnly(20, 0);

        private List<EmployeeListItemModel> Employees = new();
        private EmployeeListItemModel? SelectedEmployee { get; set; }

        private string EmployeePicBase64 { get; set; } = string.Empty;
        private List<EmpTimeDetailRow> AttendanceDetails = new();
        private string ValidationWarning { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                Employees = await AttendanceService.GetActiveEmployeesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
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

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            await LoadAttendanceForEmployeeAsync();
        }

        private async Task OnAttendanceDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime parsed))
            {
                AttendanceDate = parsed.Date;
                await LoadAttendanceForEmployeeAsync();
            }
        }

        private async Task LoadAttendanceForEmployeeAsync()
        {
            ValidationWarning = string.Empty;
            EmployeePicBase64 = string.Empty;
            AttendanceDetails.Clear();

            if (SelectedEmployee == null)
            {
                await InvokeAsync(StateHasChanged);
                return;
            }

            IsLoadingDetails = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                string empId = SelectedEmployee.EmpID;

                // 1. Check leave or holiday
                string? warn = await AttendanceService.CheckLeaveOrHolidayAsync(empId, AttendanceDate);
                if (!string.IsNullOrEmpty(warn))
                {
                    ValidationWarning = warn;
                }

                // 2. Load Photo
                var picBytes = await AttendanceService.GetEmployeePictureAsync(empId);
                if (picBytes != null && picBytes.Length > 0)
                {
                    EmployeePicBase64 = Convert.ToBase64String(picBytes);
                }

                // 3. Load Shift Default Timings
                var defaults = await AttendanceService.GetEmployeeDefaultTimingsAsync(empId, AttendanceDate);
                InTime = ParseTimeOnly(defaults.EnterTime);
                BrkOut = ParseTimeOnly(defaults.LunchInTime);
                BrkIn = ParseTimeOnly(defaults.LunchOutTime);
                ChkOut = ParseTimeOnly(defaults.ExitTime);

                // 4. Load Existing Attendance Logs
                AttendanceDetails = await AttendanceService.GetEmployeeAttendanceDetailsAsync(empId, AttendanceDate);

                // Populate from existing logs if present
                if (AttendanceDetails.Any())
                {
                    var slot1 = AttendanceDetails.FirstOrDefault(x => !x.OverTime);
                    if (slot1 != null)
                    {
                        InTime = ParseTimeOnly(slot1.InTime);
                        BrkOut = ParseTimeOnly(slot1.OutTime);
                    }

                    var slot2 = AttendanceDetails.Skip(1).FirstOrDefault(x => !x.OverTime);
                    if (slot2 != null)
                    {
                        BrkIn = ParseTimeOnly(slot2.InTime);
                        ChkOut = ParseTimeOnly(slot2.OutTime);
                    }

                    var slotOT = AttendanceDetails.FirstOrDefault(x => x.OverTime);
                    if (slotOT != null)
                    {
                        EnableOT = true;
                        OTIn = ParseTimeOnly(slotOT.InTime);
                        OTOut = ParseTimeOnly(slotOT.OutTime);
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Attendance", ex.Message);
            }
            finally
            {
                IsLoadingDetails = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task SaveAttendanceAsync()
        {
            if (SelectedEmployee == null)
            {
                NotificationService.ShowWarning("Warning", "Please select an employee.");
                return;
            }

            IsSaving = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                string clientIp = AuditService.GetClientIpAddress();

                var input = new ManualAttendanceInputDto
                {
                    EmpID = SelectedEmployee.EmpID,
                    DT = AttendanceDate,
                    UseServerTime = UseServerTime,
                    InTime = FormatTimeOnly(InTime),
                    BrkOut = FormatTimeOnly(BrkOut),
                    BrkIn = FormatTimeOnly(BrkIn),
                    ChkOut = FormatTimeOnly(ChkOut),
                    EnableOT = EnableOT,
                    OTIn = FormatTimeOnly(OTIn),
                    OTOut = FormatTimeOnly(OTOut),
                    UserName = CurrentUserName,
                    ComputerName = clientIp
                };

                bool success = await AttendanceService.SaveManualAttendanceAsync(input);
                if (success)
                {
                    NotificationService.ShowSuccess("Success", $"Attendance saved successfully for {SelectedEmployee.Name}.");
                    await LoadAttendanceForEmployeeAsync();
                }
                else
                {
                    NotificationService.ShowError("Error", "Failed to save attendance.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task DeleteAttendanceAsync()
        {
            if (SelectedEmployee == null || !AttendanceDetails.Any()) return;

            IsSaving = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                bool success = await AttendanceService.DeleteAttendanceAsync(SelectedEmployee.EmpID, AttendanceDate);
                if (success)
                {
                    NotificationService.ShowSuccess("Deleted", $"Attendance deleted for {SelectedEmployee.Name} on {AttendanceDate:dd-MMM-yyyy}.");
                    await LoadAttendanceForEmployeeAsync();
                }
                else
                {
                    NotificationService.ShowError("Error", "Failed to delete attendance.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Delete Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private void ClearForm()
        {
            InTime = new TimeOnly(8, 0);
            BrkOut = new TimeOnly(13, 0);
            BrkIn = new TimeOnly(14, 0);
            ChkOut = new TimeOnly(17, 0);
            EnableOT = false;
            OTIn = new TimeOnly(17, 30);
            OTOut = new TimeOnly(20, 0);
        }

        private TimeOnly? ParseTimeOnly(string? str)
        {
            if (string.IsNullOrWhiteSpace(str)) return null;
            if (TimeOnly.TryParse(str, out TimeOnly to)) return to;
            if (DateTime.TryParse(str, out DateTime dt)) return TimeOnly.FromDateTime(dt);
            return null;
        }

        private string FormatTimeOnly(TimeOnly? to)
        {
            return to.HasValue ? to.Value.ToString("HH:mm") : string.Empty;
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

        private bool ShowReCalcModal = false;

        private void OpenReCalcModal()
        {
            ShowReCalcModal = true;
        }

        private async Task HandleReCalculateSubmit(EmpIdAndDateRangeSelection selection)
        {
            try
            {
                int count = await AttendanceService.ReCalculateAttendanceAsync(selection.DateFrom, selection.DateTo, selection.DeptID, selection.EmpID);
                NotificationService.ShowSuccess("Re-Calculation Complete", $"Successfully re-calculated attendance for {count} record(s).");

                if (SelectedEmployee != null)
                {
                    await LoadAttendanceForEmployeeAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Re-Calculation Error", ex.Message);
            }
        }
    }
}
