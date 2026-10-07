using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.IntraOffice;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services.WorkflowTasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.Accounts
{
    public partial class VoucherApproval : ComponentBase
    {
        [Inject]
        private IWorkflowTaskEngine WorkflowTaskEngine { get; set; } = default!;

        [Inject]
        private IAccountReportingAccess AccountReportingAccess { get; set; } = default!;

        [Inject]
        private IUserRoleDataAccess UserRoleDataAccess { get; set; } = default!;

        [Inject]
        private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

        [Inject]
        private NotificationService NotificationService { get; set; } = default!;

        [Inject]
        private NavigationManager Navigation { get; set; } = default!;

        [Parameter]
        public string? pVchrNo { get; set; }

        [Parameter]
        [SupplyParameterFromQuery(Name = "vchrNo")]
        public string? QueryVchrNo { get; set; }

        [Parameter]
        [SupplyParameterFromQuery(Name = "taskId")]
        public int? QueryTaskId { get; set; }

        protected bool IsLoading { get; set; } = true;
        protected bool IsProcessing { get; set; } = false;
        protected bool IsDirector { get; set; } = false;
        protected string CurrentUserName { get; set; } = string.Empty;
        protected string? ErrorMessage { get; set; }
        protected string? CurrentVchrNo { get; set; }
        protected string DirectorRemarks { get; set; } = string.Empty;
        protected string SearchText { get; set; } = string.Empty;

        protected TaskItem? ActiveTask { get; set; }
        protected AccountsReportingModel? SelectedVoucherHeader { get; set; }
        protected List<AccountsReportingModel> SelectedVoucherData { get; set; } = new List<AccountsReportingModel>();
        protected List<TaskItem> PendingTasks { get; set; } = new List<TaskItem>();

        protected decimal TotalDebit => SelectedVoucherData.Sum(x => x.Debit);
        protected decimal TotalCredit => SelectedVoucherData.Sum(x => x.Credit);

        protected List<TaskItem> FilteredPendingTasks
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SearchText))
                    return PendingTasks;

                return PendingTasks.Where(t =>
                    (!string.IsNullOrEmpty(t.SourceEntityRefId) && t.SourceEntityRefId.Contains(SearchText, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(t.AssignedBy) && t.AssignedBy.Contains(SearchText, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(t.Description) && t.Description.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoading = true;
                await ResolveCurrentUserAndRolesAsync();
                await LoadPendingRequestsAsync();

                string? targetVchr = !string.IsNullOrEmpty(pVchrNo) ? pVchrNo : QueryVchrNo;
                if (!string.IsNullOrEmpty(targetVchr))
                {
                    await SelectVoucherForReview(targetVchr);
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Failed to initialize voucher approvals: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected override async Task OnParametersSetAsync()
        {
            string? targetVchr = !string.IsNullOrEmpty(pVchrNo) ? pVchrNo : QueryVchrNo;
            if (!string.IsNullOrEmpty(targetVchr) && targetVchr != CurrentVchrNo)
            {
                await SelectVoucherForReview(targetVchr);
            }
        }

        private async Task ResolveCurrentUserAndRolesAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;

            CurrentUserName = user.Identity?.Name ?? "Unknown";

            if (user.Identity?.IsAuthenticated == true)
            {
                // Check Director role membership
                var directorUsers = await UserRoleDataAccess.GetUsersByRoleAsync("Director");
                IsDirector = directorUsers.Any(u => string.Equals(u.UserName, CurrentUserName, StringComparison.OrdinalIgnoreCase)) ||
                             user.IsInRole("Director") ||
                             user.IsInRole("Admin") ||
                             user.IsInRole("Administrator");
            }
        }

        protected async Task LoadPendingRequestsAsync()
        {
            try
            {
                IsLoading = true;
                ErrorMessage = null;
                PendingTasks = await WorkflowTaskEngine.GetPendingVoucherDeletionTasksAsync();
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

        protected async Task SelectVoucherForReview(string vchrNo)
        {
            if (string.IsNullOrWhiteSpace(vchrNo))
                return;

            try
            {
                IsLoading = true;
                ErrorMessage = null;
                CurrentVchrNo = vchrNo.Trim();
                DirectorRemarks = string.Empty;

                // 1. Fetch voucher ledger lines
                var lines = await AccountReportingAccess.GetVoucherData(CurrentVchrNo);
                SelectedVoucherData = lines?.ToList() ?? new List<AccountsReportingModel>();

                if (SelectedVoucherData.Count > 0)
                {
                    SelectedVoucherHeader = SelectedVoucherData.First();
                }

                // 2. Locate active or completed task
                ActiveTask = PendingTasks.FirstOrDefault(t => string.Equals(t.SourceEntityRefId, CurrentVchrNo, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Failed to load details for Voucher #{vchrNo}: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected void ClearSelection()
        {
            CurrentVchrNo = null;
            ActiveTask = null;
            SelectedVoucherHeader = null;
            SelectedVoucherData.Clear();
            DirectorRemarks = string.Empty;
        }

        protected string GetFormattedDeletionReason()
        {
            if (ActiveTask != null && !string.IsNullOrWhiteSpace(ActiveTask.Description))
            {
                return ActiveTask.Description;
            }
            if (SelectedVoucherHeader != null && !string.IsNullOrWhiteSpace(SelectedVoucherHeader.DeleteReason))
            {
                return SelectedVoucherHeader.DeleteReason;
            }
            return "No reason specified.";
        }

        protected async Task ApproveAndDeleteAsync()
        {
            if (string.IsNullOrEmpty(CurrentVchrNo))
                return;

            try
            {
                IsProcessing = true;
                ErrorMessage = null;

                bool result = await WorkflowTaskEngine.ApproveVoucherDeletionAsync(
                    CurrentVchrNo, 
                    CurrentUserName, 
                    DirectorRemarks);

                if (result)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Voucher Deleted", 
                        $"Voucher #{CurrentVchrNo} deletion approved and executed. Originator has been notified.");

                    ClearSelection();
                    await LoadPendingRequestsAsync();
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Approval Failed", 
                        "Could not process voucher approval. Please try again.");
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Approval error: {ex.Message}";
                NotificationService.Notify(NotificationSeverity.Error, "Error", ex.Message);
            }
            finally
            {
                IsProcessing = false;
            }
        }

        protected async Task RejectRequestAsync()
        {
            if (string.IsNullOrEmpty(CurrentVchrNo))
                return;

            if (string.IsNullOrWhiteSpace(DirectorRemarks))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Remarks Required", 
                    "Please provide rejection remarks/reason to inform the originator.");
                return;
            }

            try
            {
                IsProcessing = true;
                ErrorMessage = null;

                bool result = await WorkflowTaskEngine.RejectVoucherDeletionAsync(
                    CurrentVchrNo, 
                    CurrentUserName, 
                    DirectorRemarks);

                if (result)
                {
                    NotificationService.Notify(NotificationSeverity.Info, "Request Rejected", 
                        $"Voucher #{CurrentVchrNo} deletion request rejected. Originator has been notified.");

                    ClearSelection();
                    await LoadPendingRequestsAsync();
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Rejection Failed", 
                        "Could not reject voucher task. Please try again.");
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Rejection error: {ex.Message}";
                NotificationService.Notify(NotificationSeverity.Error, "Error", ex.Message);
            }
            finally
            {
                IsProcessing = false;
            }
        }
    }
}
