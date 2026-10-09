using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.WorkflowTasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.Production.LotDeletionApproval
{
    public partial class LotDeletionApproval : ComponentBase
    {
        [Inject] private IWorkflowTaskEngine WorkflowTaskEngine { get; set; } = default!;
        [Inject] private IProductionDeletionDataAccess ProdDeletionDataAccess { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter] public int? pTaskId { get; set; }

        public List<ProductionDeletionRequestModel> PendingRequests { get; set; } = new List<ProductionDeletionRequestModel>();
        public ProductionDeletionRequestModel? SelectedRequest { get; set; } = null;
        public string DirectorRemarks { get; set; } = string.Empty;
        public bool IsLoading { get; set; } = false;
        public bool IsProcessingAction { get; set; } = false;
        public string? ErrorMessage { get; set; }

        private string CurrentUserName { get; set; } = "Director";

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                CurrentUserName = authState.User.Identity?.Name ?? "Director";
            }
            catch
            {
                CurrentUserName = "Director";
            }

            await LoadPendingRequestsAsync();

            if (pTaskId.HasValue && pTaskId.Value > 0)
            {
                await LoadByTaskIdAsync(pTaskId.Value);
            }
        }

        public async Task LoadPendingRequestsAsync()
        {
            IsLoading = true;
            ErrorMessage = null;
            StateHasChanged();

            try
            {
                PendingRequests = await WorkflowTaskEngine.GetPendingLotDeletionRequestsAsync();
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Failed to initialize lot deletion approval queue: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private async Task LoadByTaskIdAsync(int taskId)
        {
            try
            {
                var match = PendingRequests.FirstOrDefault(r => r.TaskId == taskId);
                if (match != null)
                {
                    SelectRequest(match);
                    return;
                }

                // If not found in pending, fetch directly
                var all = await ProdDeletionDataAccess.GetPendingRequestsAsync();
                var directMatch = all.FirstOrDefault(r => r.TaskId == taskId);
                if (directMatch != null)
                {
                    SelectRequest(directMatch);
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Error loading task #{taskId}: {ex.Message}";
            }
        }

        public void SelectRequest(ProductionDeletionRequestModel request)
        {
            SelectedRequest = request;
            DirectorRemarks = string.Empty;
            ErrorMessage = null;
        }

        public void ClearSelection()
        {
            SelectedRequest = null;
            DirectorRemarks = string.Empty;
            ErrorMessage = null;
        }

        public async Task ApproveRequestAsync()
        {
            if (SelectedRequest == null) return;

            IsProcessingAction = true;
            ErrorMessage = null;
            StateHasChanged();

            try
            {
                bool success = await WorkflowTaskEngine.ApproveLotReceivingDeletionAsync(
                    SelectedRequest.Id,
                    CurrentUserName,
                    string.IsNullOrWhiteSpace(DirectorRemarks) ? "Approved & Deleted" : DirectorRemarks.Trim());

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Lot Deleted",
                        Detail = $"Lot #{SelectedRequest.LotNo} receiving has been approved and deleted successfully.",
                        Duration = 5000
                    });

                    SelectedRequest = null;
                    await LoadPendingRequestsAsync();
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Approval Failed: {ex.Message}";
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Approval Blocked",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsProcessingAction = false;
                StateHasChanged();
            }
        }

        public async Task RejectRequestAsync()
        {
            if (SelectedRequest == null) return;

            if (string.IsNullOrWhiteSpace(DirectorRemarks))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Remarks Required",
                    Detail = "Please provide a reason in the remarks field before rejecting this request.",
                    Duration = 4000
                });
                return;
            }

            IsProcessingAction = true;
            ErrorMessage = null;
            StateHasChanged();

            try
            {
                bool success = await WorkflowTaskEngine.RejectLotReceivingDeletionAsync(
                    SelectedRequest.Id,
                    CurrentUserName,
                    DirectorRemarks.Trim());

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Info,
                        Summary = "Request Rejected",
                        Detail = $"Deletion request for Lot #{SelectedRequest.LotNo} was rejected.",
                        Duration = 4000
                    });

                    SelectedRequest = null;
                    await LoadPendingRequestsAsync();
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Rejection Failed: {ex.Message}";
            }
            finally
            {
                IsProcessingAction = false;
                StateHasChanged();
            }
        }
    }
}
