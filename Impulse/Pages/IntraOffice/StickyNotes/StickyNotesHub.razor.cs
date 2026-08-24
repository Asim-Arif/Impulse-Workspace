using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.IntraOffice.StickyNotes
{
    public partial class StickyNotesHub : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;

        protected bool IsLoading { get; set; } = true;
        protected string CurrentUserId { get; set; } = string.Empty;

        protected List<StickyNote> NotesList { get; set; } = new();

        protected readonly string[] AvailableColors = new[]
        {
            "#fff9c4", // Yellow
            "#c8e6c9", // Light Green
            "#b3e5fc", // Light Blue
            "#ffcdd2", // Light Pink
            "#e1bee7"  // Light Purple
        };

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
                NotesList = await IntraOfficeService.GetStickyNotesAsync(CurrentUserId);
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

        protected async Task CreateNewNote()
        {
            try
            {
                var note = new StickyNote
                {
                    UserId = CurrentUserId,
                    Content = "New note...",
                    Color = AvailableColors[new Random().Next(AvailableColors.Length)],
                    CreatedAt = DateTime.UtcNow
                };

                var id = await IntraOfficeService.SaveStickyNoteAsync(note);
                if (id > 0)
                {
                    note.Id = id;
                    NotesList.Insert(0, note);
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

        protected async Task SaveNoteAsync(StickyNote note)
        {
            try
            {
                await IntraOfficeService.SaveStickyNoteAsync(note);
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
        }

        protected async Task ChangeNoteColor(StickyNote note, string color)
        {
            note.Color = color;
            await SaveNoteAsync(note);
        }

        protected async Task DeleteNoteAsync(int noteId)
        {
            try
            {
                var success = await IntraOfficeService.DeleteStickyNoteAsync(noteId, CurrentUserId);
                if (success)
                {
                    NotesList.RemoveAll(n => n.Id == noteId);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Delete Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }
    }
}
