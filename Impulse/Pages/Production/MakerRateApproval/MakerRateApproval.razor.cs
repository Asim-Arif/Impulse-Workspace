using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using DataAccessLibrary.Models.IntraOffice;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.WorkflowTasks;

namespace Impulse.Pages.Production.MakerRateApproval
{
    public partial class MakerRateApproval : ComponentBase
    {
        [Parameter] public int? pTaskId { get; set; }

        [SupplyParameterFromQuery(Name = "taskId")]
        public int? QueryTaskId { get; set; }

        [Inject] private IWorkflowTaskEngine WorkflowTaskEngine { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        private bool IsLoading = false;
        private bool IsProcessing = false;
        private string? ErrorMessage = null;
        private string CurrentUserName = "System";

        private List<TaskItem> PendingTasks = new();
        private int SelectedTaskId = 0;
        private TaskItem? ActiveTask = null;
        private MakerItemRateWorkflowRequestDto? SelectedRequest = null;

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
                ErrorMessage = $"Failed to initialize maker rate approval queue: {ex.Message}";
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
                PendingTasks = await WorkflowTaskEngine.GetPendingMakerRateTasksAsync();
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
                SelectedRequest = await WorkflowTaskEngine.GetMakerRateTaskDetailsAsync(taskId);
                ActiveTask = PendingTasks.Find(t => t.Id == taskId);

                if (SelectedRequest == null)
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
        }

        private async Task ApproveRequestAsync()
        {
            if (SelectedRequest == null || SelectedTaskId <= 0) return;

            IsProcessing = true;
            ErrorMessage = null;

            try
            {
                bool success = await WorkflowTaskEngine.ApproveMakerItemRateActionAsync(SelectedTaskId, CurrentUserName, DirectorRemarks);
                if (success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Approved & Applied",
                        Detail = $"Rate change for [{SelectedRequest.ItemID}] ({SelectedRequest.MakerName}) was approved and applied at {SelectedRequest.NewRate:N2}.",
                        Duration = 5000
                    });

                    ClearSelection();
                    await LoadPendingRequestsAsync();
                }
                else
                {
                    ErrorMessage = "Failed to approve rate change.";
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Approval Failed: {ex.Message}";
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

        private void CloseRejectModal()
        {
            ShowRejectModal = false;
            RejectReasonText = string.Empty;
        }

        private async Task ConfirmRejectAsync()
        {
            if (SelectedRequest == null || SelectedTaskId <= 0) return;
            if (string.IsNullOrWhiteSpace(RejectReasonText))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Reason Required",
                    Detail = "Please provide a rejection reason.",
                    Duration = 4000
                });
                return;
            }

            IsProcessing = true;
            ErrorMessage = null;

            try
            {
                bool success = await WorkflowTaskEngine.RejectMakerItemRateActionAsync(SelectedTaskId, CurrentUserName, RejectReasonText.Trim());
                if (success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Info,
                        Summary = "Request Rejected",
                        Detail = $"Rate change request for [{SelectedRequest.ItemID}] was rejected.",
                        Duration = 5000
                    });

                    CloseRejectModal();
                    ClearSelection();
                    await LoadPendingRequestsAsync();
                }
                else
                {
                    ErrorMessage = "Failed to reject rate change request.";
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Rejection Failed: {ex.Message}";
            }
            finally
            {
                IsProcessing = false;
            }
        }
    }
}
