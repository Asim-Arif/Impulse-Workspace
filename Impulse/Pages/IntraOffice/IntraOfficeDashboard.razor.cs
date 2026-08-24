using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;

namespace Impulse.Pages.IntraOffice
{
    public partial class IntraOfficeDashboard : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;

        protected string CurrentUserId { get; set; } = string.Empty;

        protected int UnreadMessagesCount { get; set; } = 0;
        protected int ActiveChannelsCount { get; set; } = 0;
        protected int AnnouncementsCount { get; set; } = 0;
        protected int PendingTasksCount { get; set; } = 0;
        protected int UpcomingMeetingsCount { get; set; } = 0;
        protected int PendingMinutesCount { get; set; } = 0;

        protected List<Announcement> RecentAnnouncements { get; set; } = new();
        protected List<TaskItem> MyTasks { get; set; } = new();
        protected List<Meeting> UpcomingMeetings { get; set; } = new();
        protected List<StickyNote> RemindersList { get; set; } = new();

        protected bool ShowAddReminder { get; set; } = false;
        protected string NewReminderContent { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;
            CurrentUserId = user.Identity?.Name ?? user.FindFirst(ClaimTypes.Name)?.Value ?? "Guest";

            await LoadDashboardDataAsync();
        }

        protected async Task LoadDashboardDataAsync()
        {
            try
            {
                var directSummaries = await IntraOfficeService.GetDirectChatSummariesAsync(CurrentUserId);
                UnreadMessagesCount = directSummaries.Sum(s => s.UnreadCount);

                var channels = await IntraOfficeService.GetChannelsForUserAsync(CurrentUserId);
                ActiveChannelsCount = channels.Count;

                var announcements = await IntraOfficeService.GetAnnouncementsAsync(CurrentUserId);
                AnnouncementsCount = announcements.Count;
                RecentAnnouncements = announcements;

                var tasks = await IntraOfficeService.GetTasksAsync(assignedTo: CurrentUserId);
                MyTasks = tasks;
                PendingTasksCount = tasks.Count(t => t.Status == TaskItemStatus.Pending || t.Status == TaskItemStatus.InProgress);

                var meetings = await IntraOfficeService.GetMeetingsForUserAsync(CurrentUserId);
                UpcomingMeetings = meetings;
                UpcomingMeetingsCount = meetings.Count;

                var minutes = await IntraOfficeService.GetMinutesAsync(new MinuteFilter { Status = "Pending" });
                PendingMinutesCount = minutes.Count;

                var notes = await IntraOfficeService.GetStickyNotesAsync(CurrentUserId);
                RemindersList = notes;
            }
            catch (Exception)
            {
                // Fallback safe
            }
        }

        protected async Task AddReminderAsync()
        {
            if (string.IsNullOrWhiteSpace(NewReminderContent)) return;

            try
            {
                var note = new StickyNote
                {
                    UserId = CurrentUserId,
                    Content = NewReminderContent.Trim(),
                    Color = "yellow"
                };

                await IntraOfficeService.SaveStickyNoteAsync(note);
                NewReminderContent = string.Empty;
                ShowAddReminder = false;
                RemindersList = await IntraOfficeService.GetStickyNotesAsync(CurrentUserId);
            }
            catch
            {
                // Graceful
            }
        }

        protected async Task DeleteReminderAsync(int noteId)
        {
            try
            {
                await IntraOfficeService.DeleteStickyNoteAsync(noteId, CurrentUserId);
                RemindersList = await IntraOfficeService.GetStickyNotesAsync(CurrentUserId);
            }
            catch
            {
                // Graceful
            }
        }
    }
}
