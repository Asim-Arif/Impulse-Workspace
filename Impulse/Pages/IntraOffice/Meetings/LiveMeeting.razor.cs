using System;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.IntraOffice.Meetings
{
    public partial class LiveMeeting : ComponentBase
    {
        [Parameter] public int MeetingId { get; set; }

        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;

        protected Meeting? MeetingModel { get; set; }
        protected string CurrentUserId { get; set; } = string.Empty;
        protected string MeetingNotes { get; set; } = string.Empty;
        protected bool IsSavingNotes { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;
            CurrentUserId = user.Identity?.Name ?? user.FindFirst(ClaimTypes.Name)?.Value ?? "Guest";

            await LoadMeetingAsync();
        }

        protected async Task LoadMeetingAsync()
        {
            try
            {
                var meetings = await IntraOfficeService.GetMeetingsForUserAsync(CurrentUserId);
                MeetingModel = await IntraOfficeService.GetMeetingByIdAsync(MeetingId);

                if (MeetingModel != null)
                {
                    MeetingNotes = MeetingModel.MeetingMinutes ?? string.Empty;
                    if (string.IsNullOrEmpty(MeetingModel.MeetingUrl))
                    {
                        MeetingModel.MeetingUrl = $"https://meet.jit.si/Impulse_Meeting_{MeetingModel.Id}_{Guid.NewGuid():N}";
                    }
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Not Found",
                        Detail = "The requested meeting room was not found.",
                        Duration = 4000
                    });
                    NavigationManager.NavigateTo("/intraoffice/meetings");
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

        protected async Task SaveMeetingNotesAsync()
        {
            if (MeetingModel == null) return;

            IsSavingNotes = true;
            try
            {
                MeetingModel.MeetingMinutes = MeetingNotes;
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Saved",
                    Detail = "Meeting notes updated successfully.",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSavingNotes = false;
            }
        }
    }
}
