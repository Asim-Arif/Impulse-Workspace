using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.Setups.RemoteAccess
{
    public partial class RemoteAccessRequests : ComponentBase
    {
        [Inject] private IRemoteAccessDataAccess RemoteAccessData { get; set; } = default!;
        [Inject] private IUserDataAccess UserData { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;

        protected List<RemoteAccessRequestModel> PendingRequests { get; set; } = new();
        protected List<RemoteAccessRequestModel> HistoryRequests { get; set; } = new();
        protected Dictionary<int, int> SelectedDurations { get; set; } = new();

        protected int PendingCount => PendingRequests.Count;
        protected int ActiveTab { get; set; } = 0;
        protected bool IsLoading { get; set; } = true;

        protected string CurrentUserName { get; set; } = "Admin";
        protected int CurrentUserId { get; set; } = 1;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                var user = authState.User;
                if (user?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = user.Identity.Name ?? "Admin";
                    var idClaim = user.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                    if (int.TryParse(idClaim, out int uid))
                    {
                        CurrentUserId = uid;
                    }
                    else
                    {
                        var dbUser = await UserData.GetUserByUserNameAsync(CurrentUserName);
                        if (dbUser != null) CurrentUserId = dbUser.UserID;
                    }
                }

                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Initialization Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        protected async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                PendingRequests = await RemoteAccessData.GetPendingRequestsAsync();
                HistoryRequests = await RemoteAccessData.GetRecentRequestsAsync(100);

                foreach (var req in PendingRequests)
                {
                    if (!SelectedDurations.ContainsKey(req.EntryID))
                    {
                        SelectedDurations[req.EntryID] = 60; // Default 1 Hour
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to load requests",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        protected void SetTab(int tab)
        {
            ActiveTab = tab;
            StateHasChanged();
        }

        protected async Task ApproveAsync(int entryId)
        {
            int duration = SelectedDurations.TryGetValue(entryId, out int d) ? d : 60;
            var req = PendingRequests.FirstOrDefault(r => r.EntryID == entryId);

            try
            {
                var success = await RemoteAccessData.ApproveRequestAsync(
                    entryId, 
                    CurrentUserId, 
                    CurrentUserName, 
                    duration, 
                    $"Approved for {duration} minutes"
                );

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Access Granted",
                        Detail = $"User {req?.UserName ?? "Staff"} can now work remotely for {duration} minutes.",
                        Duration = 5000
                    });

                    await LoadDataAsync();
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Approval Failed",
                        Detail = "This request could not be approved or was already processed.",
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        protected async Task RejectAsync(int entryId)
        {
            var req = PendingRequests.FirstOrDefault(r => r.EntryID == entryId);
            try
            {
                var success = await RemoteAccessData.RejectRequestAsync(
                    entryId, 
                    CurrentUserId, 
                    CurrentUserName, 
                    "Rejected by administrator"
                );

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Info,
                        Summary = "Request Rejected",
                        Detail = $"Remote access request for {req?.UserName ?? "Staff"} was rejected.",
                        Duration = 4000
                    });

                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }
    }
}

