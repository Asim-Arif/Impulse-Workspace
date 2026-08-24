using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.IntraOffice.Meetings
{
    public partial class MeetingsList : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;

        protected bool IsLoading { get; set; } = true;
        protected string CurrentUserId { get; set; } = string.Empty;

        protected List<Meeting> AllMeetings { get; set; } = new();
        protected List<IntraUserProfile> UsersList { get; set; } = new();

        // Modal fields
        protected bool ShowCreateModalState { get; set; } = false;
        protected string NewMeetingTitle { get; set; } = string.Empty;
        protected DateTime NewMeetingStartTime { get; set; } = DateTime.Now.AddHours(1);
        protected string NewMeetingUrl { get; set; } = string.Empty;
        protected string NewMeetingMinutes { get; set; } = string.Empty;
        protected List<string> SelectedParticipantIds { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;
            CurrentUserId = user.Identity?.Name ?? user.FindFirst(ClaimTypes.Name)?.Value ?? "Guest";

            await LoadDataAsync();
        }

        protected async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                UsersList = await IntraOfficeService.GetActiveUsersAsync();
                AllMeetings = await IntraOfficeService.GetMeetingsForUserAsync(CurrentUserId);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Load Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected void OpenCreateModal()
        {
            NewMeetingTitle = string.Empty;
            NewMeetingStartTime = DateTime.Now.AddHours(1);
            var roomCode = Guid.NewGuid().ToString("N")[..8];
            NewMeetingUrl = $"https://meet.jit.si/impulse-{roomCode}";
            NewMeetingMinutes = string.Empty;
            SelectedParticipantIds = new List<string>();
            ShowCreateModalState = true;
        }

        protected void CloseCreateModal()
        {
            ShowCreateModalState = false;
        }

        protected void ToggleParticipant(string userId, bool isSelected)
        {
            if (isSelected && !SelectedParticipantIds.Contains(userId))
            {
                SelectedParticipantIds.Add(userId);
            }
            else if (!isSelected)
            {
                SelectedParticipantIds.Remove(userId);
            }
        }

        protected async Task SaveMeetingAsync()
        {
            if (string.IsNullOrWhiteSpace(NewMeetingTitle))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter a meeting title.",
                    Duration = 3000
                });
                return;
            }

            try
            {
                var meeting = new Meeting
                {
                    Title = NewMeetingTitle.Trim(),
                    OrganizerId = CurrentUserId,
                    ScheduledStartTime = NewMeetingStartTime,
                    MeetingUrl = string.IsNullOrWhiteSpace(NewMeetingUrl) ? null : NewMeetingUrl.Trim(),
                    MeetingMinutes = string.IsNullOrWhiteSpace(NewMeetingMinutes) ? null : NewMeetingMinutes.Trim(),
                    Status = MeetingStatus.Scheduled
                };

                var id = await IntraOfficeService.CreateMeetingAsync(meeting, SelectedParticipantIds);
                if (id > 0)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Scheduled",
                        Detail = "Meeting scheduled successfully.",
                        Duration = 3000
                    });

                    ShowCreateModalState = false;
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

        protected async Task CancelMeetingAsync(int meetingId)
        {
            try
            {
                var success = await IntraOfficeService.UpdateMeetingStatusAsync(meetingId, MeetingStatus.Cancelled);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Info,
                        Summary = "Cancelled",
                        Detail = "Meeting was cancelled.",
                        Duration = 3000
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

        protected string GetStatusBadgeClass(MeetingStatus status)
        {
            return status switch
            {
                MeetingStatus.InProgress => "bg-primary",
                MeetingStatus.Completed => "bg-success",
                MeetingStatus.Cancelled => "bg-danger",
                _ => "bg-warning text-dark"
            };
        }
    }
}
