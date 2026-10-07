using DataAccessLibrary.Models.IntraOffice;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Impulse.Services.WorkflowTasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.AttendanceApproval
{
    public partial class AttendanceApproval : ComponentBase
    {
        [Parameter] public int? pTaskId { get; set; }

        [SupplyParameterFromQuery(Name = "taskId")]
        public int? QueryTaskId { get; set; }

        [Inject] private IWorkflowTaskEngine WorkflowTaskEngine { get; set; } = default!;
        [Inject] private IMonthlyAttendanceService MonthlyAttendanceService { get; set; } = default!;
        [Inject] private INotificationService NotificationService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        private bool IsLoading = false;
        private bool IsProcessing = false;
        private string? ErrorMessage = null;
        private string CurrentUserName = "System";

        private List<TaskItem> PendingTasks = new();
        private int SelectedTaskId = 0;
        private TaskItem? ActiveTask = null;
        private AttendanceWorkflowRequestDto? SelectedRequest = null;
        private bool IsSalaryFinalizedForSelected = false;

        private string DirectorRemarks = string.Empty;
        private bool ShowRejectModal = false;
        private string RejectReasonText = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                await LoadPendingRequestsAsync();

                int targetId = pTaskId ?? QueryTaskId ?? 0;
                if (targetId > 0)
                {
                    await SelectTaskAsync(targetId);
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Failed to initialize attendance approval queue: {ex.Message}";
            }
        }

        protected override async Task OnParametersSetAsync()
        {
            int targetId = pTaskId ?? QueryTaskId ?? 0;
            if (targetId > 0 && targetId != SelectedTaskId)
            {
                await SelectTaskAsync(targetId);
            }
        }

        private async Task LoadPendingRequestsAsync()
        {
            IsLoading = true;
            ErrorMessage = null;
            try
            {
                PendingTasks = await WorkflowTaskEngine.GetPendingAttendanceTasksAsync();
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Error loading pending requests: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task SelectTaskAsync(int taskId)
        {
            if (taskId <= 0) return;

            IsLoading = true;
            ErrorMessage = null;
            SelectedTaskId = taskId;
            DirectorRemarks = string.Empty;

            try
            {
                SelectedRequest = await WorkflowTaskEngine.GetAttendanceTaskDetailsAsync(taskId);
                ActiveTask = PendingTasks.Find(t => t.Id == taskId);

                if (SelectedRequest != null)
                {
                    // Check month-wide salary finalization
                    IsSalaryFinalizedForSelected = await MonthlyAttendanceService.IsSalaryFinalizedAsync(SelectedRequest.Year, SelectedRequest.Month);
                }
                else
                {
                    ErrorMessage = $"Could not load details for task #{taskId}.";
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Error loading task #{taskId}: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ClearSelection()
        {
            SelectedTaskId = 0;
            SelectedRequest = null;
            ActiveTask = null;
            DirectorRemarks = string.Empty;
            IsSalaryFinalizedForSelected = false;
        }

        private async Task ApproveRequestAsync()
        {
            if (SelectedRequest == null || SelectedTaskId <= 0) return;

            IsProcessing = true;
            ErrorMessage = null;

            try
            {
                // Re-verify salary finalization before executing
                bool isFinal = await MonthlyAttendanceService.IsSalaryFinalizedAsync(SelectedRequest.Year, SelectedRequest.Month);
                if (isFinal)
                {
                    IsSalaryFinalizedForSelected = true;
                    NotificationService.ShowError("Approval Blocked", $"Salary for {SelectedRequest.Year}-{SelectedRequest.Month:D2} has already been finalized in MonthlySalaries.");
                    return;
                }

                bool success = await WorkflowTaskEngine.ApproveAttendanceActionAsync(SelectedTaskId, CurrentUserName, DirectorRemarks);
                if (success)
                {
                    NotificationService.ShowSuccess("Approved & Applied", $"Attendance modification request for [{SelectedRequest.EmpID}] {SelectedRequest.EmployeeName} was approved and applied successfully.");
                    ClearSelection();
                    await LoadPendingRequestsAsync();
                }
                else
                {
                    NotificationService.ShowError("Approval Failed", "Failed to approve attendance request.");
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Approval error: {ex.Message}";
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                IsProcessing = false;
            }
        }

        private void OpenRejectModal()
        {
            RejectReasonText = string.Empty;
            ShowRejectModal = true;
        }

        private async Task ConfirmRejectAsync()
        {
            if (SelectedRequest == null || SelectedTaskId <= 0) return;
            if (string.IsNullOrWhiteSpace(RejectReasonText))
            {
                NotificationService.ShowWarning("Reason Required", "Please enter rejection remarks.");
                return;
            }

            IsProcessing = true;
            ShowRejectModal = false;

            try
            {
                bool success = await WorkflowTaskEngine.RejectAttendanceActionAsync(SelectedTaskId, CurrentUserName, RejectReasonText);
                if (success)
                {
                    NotificationService.ShowSuccess("Rejected", $"Attendance request for [{SelectedRequest.EmpID}] {SelectedRequest.EmployeeName} was rejected.");
                    ClearSelection();
                    await LoadPendingRequestsAsync();
                }
                else
                {
                    NotificationService.ShowError("Rejection Failed", "Failed to reject attendance request.");
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Rejection error: {ex.Message}";
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                IsProcessing = false;
            }
        }
    }
}
