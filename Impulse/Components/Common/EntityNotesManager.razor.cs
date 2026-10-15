using System;
using System.Threading.Tasks;
using Impulse.Services.Common;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Components.Common
{
    public partial class EntityNotesManager : ComponentBase
    {
        [Inject] private IEntityNoteService NoteService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

        [Parameter] public string EntityType { get; set; } = string.Empty;
        [Parameter] public string EntityId { get; set; } = string.Empty;
        [Parameter] public string? EntitySecondaryId { get; set; }
        [Parameter] public bool IsReadOnly { get; set; } = false;

        private string CurrentNoteText = string.Empty;
        private string InitialNoteText = string.Empty;
        private string? LastModifiedBy;
        private DateTime? LastModifiedAt;
        private bool IsLoading = true;
        private bool IsSaving = false;
        private string _lastLoadedKey = string.Empty;

        private bool HasUnsavedChanges => CurrentNoteText != InitialNoteText;
        private int CharacterCount => CurrentNoteText?.Length ?? 0;

        protected override async Task OnParametersSetAsync()
        {
            var currentKey = $"{EntityType}_{EntityId}_{EntitySecondaryId}";
            if (currentKey != _lastLoadedKey)
            {
                _lastLoadedKey = currentKey;
                await LoadNoteAsync();
            }
        }

        public async Task LoadNoteAsync()
        {
            if (string.IsNullOrWhiteSpace(EntityType) || string.IsNullOrWhiteSpace(EntityId))
            {
                CurrentNoteText = string.Empty;
                InitialNoteText = string.Empty;
                LastModifiedBy = null;
                LastModifiedAt = null;
                IsLoading = false;
                return;
            }

            try
            {
                IsLoading = true;
                var note = await NoteService.GetNoteAsync(EntityType, EntityId, EntitySecondaryId);
                if (note != null)
                {
                    CurrentNoteText = note.NoteText ?? string.Empty;
                    InitialNoteText = CurrentNoteText;
                    LastModifiedBy = note.LastModifiedBy;
                    LastModifiedAt = note.LastModifiedAt;
                }
                else
                {
                    CurrentNoteText = string.Empty;
                    InitialNoteText = string.Empty;
                    LastModifiedBy = null;
                    LastModifiedAt = null;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Notes Error",
                    Detail = $"Could not load note: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        public async Task SaveNoteAsync()
        {
            if (string.IsNullOrWhiteSpace(EntityType) || string.IsNullOrWhiteSpace(EntityId))
            {
                return;
            }

            try
            {
                IsSaving = true;
                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                var currentUserName = authState.User.Identity?.Name ?? "Admin";

                var success = await NoteService.SaveNoteAsync(EntityType, EntityId, EntitySecondaryId, CurrentNoteText, currentUserName);
                if (success)
                {
                    InitialNoteText = CurrentNoteText;
                    LastModifiedBy = currentUserName;
                    LastModifiedAt = DateTime.Now;

                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Note Saved",
                        Detail = "Notes have been saved successfully.",
                        Duration = 3000
                    });
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Save Failed",
                        Detail = "Unable to save notes. Please try again.",
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Saving Note",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
                StateHasChanged();
            }
        }

        private async Task InsertTimestampAndUser()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var currentUserName = authState.User.Identity?.Name ?? "User";
            var headerStamp = $"--- [{DateTime.Now:dd-MMM-yyyy hh:mm tt} by {currentUserName}] ---\n";

            if (string.IsNullOrWhiteSpace(CurrentNoteText))
            {
                CurrentNoteText = headerStamp;
            }
            else
            {
                // Prepend at the top so newest stamped notes are immediately visible
                CurrentNoteText = headerStamp + "\n" + CurrentNoteText.TrimStart();
            }
            StateHasChanged();
        }

        private void OnTextChanged(ChangeEventArgs e)
        {
            CurrentNoteText = e.Value?.ToString() ?? string.Empty;
        }

        private async Task HandleKeyDown(KeyboardEventArgs e)
        {
            if (e.CtrlKey && (e.Key == "s" || e.Key == "S"))
            {
                await SaveNoteAsync();
            }
        }
    }
}
