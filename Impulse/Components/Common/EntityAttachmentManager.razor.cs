using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Attachments;
using Impulse.Services.Common;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Forms;
using Radzen;

namespace Impulse.Components.Common
{
    public partial class EntityAttachmentManager : ComponentBase
    {
        [Inject] private IEntityAttachmentService AttachmentService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

        [Parameter] public string EntityType { get; set; } = string.Empty;
        [Parameter] public string EntityId { get; set; } = string.Empty;
        [Parameter] public string? EntitySecondaryId { get; set; }
        [Parameter] public string? EntityDisplayName { get; set; }
        [Parameter] public bool IsReadOnly { get; set; } = false;
        [Parameter] public EventCallback<int> OnCountChanged { get; set; }

        private List<EntityAttachmentDto> Attachments = new();
        private List<IBrowserFile> SelectedFiles = new();
        private bool IsLoading = true;
        private bool IsUploading = false;
        private string UploadRemarks = string.Empty;
        private string? ErrorMessage;
        private string _lastLoadedKey = string.Empty;

        protected override async Task OnParametersSetAsync()
        {
            var currentKey = $"{EntityType}_{EntityId}_{EntitySecondaryId}";
            if (currentKey != _lastLoadedKey)
            {
                _lastLoadedKey = currentKey;
                await LoadAttachmentsAsync();
            }
        }

        public async Task LoadAttachmentsAsync()
        {
            if (string.IsNullOrWhiteSpace(EntityType) || string.IsNullOrWhiteSpace(EntityId))
            {
                Attachments.Clear();
                IsLoading = false;
                return;
            }

            try
            {
                IsLoading = true;
                ErrorMessage = null;
                Attachments = await AttachmentService.GetAttachmentsAsync(EntityType, EntityId, EntitySecondaryId);
                if (OnCountChanged.HasDelegate)
                {
                    await OnCountChanged.InvokeAsync(Attachments.Count);
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Failed to load attachments: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private void HandleFilesSelected(InputFileChangeEventArgs e)
        {
            SelectedFiles.Clear();
            ErrorMessage = null;

            const long maxFileSize = 50 * 1024 * 1024; // 50MB
            foreach (var file in e.GetMultipleFiles(maximumFileCount: 20))
            {
                if (file.Size > maxFileSize)
                {
                    ErrorMessage = $"File '{file.Name}' exceeds the maximum allowed size of 50MB.";
                    return;
                }
                SelectedFiles.Add(file);
            }
        }

        private async Task TriggerUploadAsync()
        {
            if (SelectedFiles.Count == 0 || string.IsNullOrWhiteSpace(EntityId))
            {
                ErrorMessage = "Please select at least one file to upload.";
                return;
            }

            try
            {
                IsUploading = true;
                ErrorMessage = null;

                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                var currentUserName = authState.User.Identity?.Name ?? "Admin";

                int uploadedCount = 0;
                foreach (var file in SelectedFiles)
                {
                    await AttachmentService.UploadAttachmentAsync(
                        entityType: EntityType,
                        entityId: EntityId,
                        entitySecondaryId: EntitySecondaryId,
                        file: file,
                        remarks: string.IsNullOrWhiteSpace(UploadRemarks) ? null : UploadRemarks,
                        uploadedBy: currentUserName
                    );
                    uploadedCount++;
                }

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Files Uploaded",
                    Detail = $"Successfully uploaded {uploadedCount} attachment(s).",
                    Duration = 4000
                });

                SelectedFiles.Clear();
                UploadRemarks = string.Empty;
                await LoadAttachmentsAsync();
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Upload failed: {ex.Message}";
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Upload Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsUploading = false;
                StateHasChanged();
            }
        }

        private async Task RequestDeleteAsync(EntityAttachmentDto item)
        {
            try
            {
                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                var currentUserName = authState.User.Identity?.Name ?? "Admin";

                var success = await AttachmentService.DeleteAttachmentAsync(item.Id, currentUserName);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Info,
                        Summary = "Attachment Removed",
                        Detail = $"'{item.OriginalFileName}' was deleted.",
                        Duration = 3000
                    });

                    await LoadAttachmentsAsync();
                }
                else
                {
                    ErrorMessage = "Failed to delete attachment.";
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Delete failed: {ex.Message}";
            }
        }

        private string FormatBytes(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{(bytes / 1024.0):F1} KB";
            return $"{(bytes / (1024.0 * 1024.0)):F2} MB";
        }
    }
}
