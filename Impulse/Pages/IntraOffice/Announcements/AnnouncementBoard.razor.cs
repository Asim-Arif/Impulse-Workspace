using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Hosting;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.IntraOffice.Announcements
{
    public partial class AnnouncementBoard : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private IWebHostEnvironment WebHostEnvironment { get; set; } = null!;
        [Inject] private IJSRuntime JS { get; set; } = null!;

        protected string CurrentUserId { get; set; } = string.Empty;
        protected bool IsLoading { get; set; } = true;
        protected bool IsPublishing { get; set; } = false;
        protected bool IsRecordingVoice { get; set; } = false;

        protected List<Announcement> AllAnnouncements { get; set; } = new();
        protected List<(string DeptId, string DeptName)> Departments { get; set; } = new();

        protected string SearchTerm { get; set; } = string.Empty;
        protected string SelectedDepartmentId { get; set; } = string.Empty;
        protected int SelectedPriorityFilter { get; set; } = -1;

        // Modal fields
        protected bool ShowCreateModalState { get; set; } = false;
        protected string NewTitle { get; set; } = string.Empty;
        protected string NewContent { get; set; } = string.Empty;
        protected string NewDepartmentId { get; set; } = string.Empty;
        protected int NewPriority { get; set; } = 0;
        protected bool NewIsPinned { get; set; } = false;

        protected List<IBrowserFile> PendingFiles { get; set; } = new();
        protected List<AnnouncementAttachment> PendingVoiceNotes { get; set; } = new();

        public class AudioRecordingResult
        {
            public string Base64 { get; set; } = string.Empty;
            public string MimeType { get; set; } = string.Empty;
        }

        protected List<Announcement> FilteredAnnouncements
        {
            get
            {
                var query = AllAnnouncements.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(SelectedDepartmentId))
                {
                    query = query.Where(a => string.Equals(a.DepartmentId, SelectedDepartmentId, StringComparison.OrdinalIgnoreCase));
                }

                if (SelectedPriorityFilter >= 0)
                {
                    query = query.Where(a => (int)a.Priority == SelectedPriorityFilter);
                }

                if (!string.IsNullOrWhiteSpace(SearchTerm))
                {
                    query = query.Where(a =>
                        a.Title.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        a.Content.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        (a.CreatorName?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false));
                }

                return query.OrderByDescending(a => a.IsPinned).ThenByDescending(a => a.CreatedAt).ToList();
            }
        }

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
                Departments = await IntraOfficeService.GetDepartmentsAsync();
                AllAnnouncements = await IntraOfficeService.GetAnnouncementsAsync(CurrentUserId);
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
            finally
            {
                IsLoading = false;
            }
        }

        protected async Task ToggleVoiceRecordingAsync()
        {
            try
            {
                if (!IsRecordingVoice)
                {
                    var success = await JS.InvokeAsync<bool>("window.audioRecorder.startRecording");
                    if (success)
                    {
                        IsRecordingVoice = true;
                    }
                    else
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Warning,
                            Summary = "Microphone Access",
                            Detail = "Please allow microphone access in your browser to record voice notes.",
                            Duration = 4000
                        });
                    }
                }
                else
                {
                    IsRecordingVoice = false;
                    var result = await JS.InvokeAsync<AudioRecordingResult>("window.audioRecorder.stopRecording");

                    if (result != null && !string.IsNullOrWhiteSpace(result.Base64))
                    {
                        var uploadFolder = System.IO.Path.Combine(WebHostEnvironment.WebRootPath ?? "wwwroot", "uploads", "announcements");
                        System.IO.Directory.CreateDirectory(uploadFolder);

                        var fileName = $"VoiceNote_{DateTime.Now:yyyyMMdd_HHmmss}.webm";
                        var filePath = System.IO.Path.Combine(uploadFolder, fileName);
                        var bytes = Convert.FromBase64String(result.Base64);

                        await System.IO.File.WriteAllBytesAsync(filePath, bytes);

                        PendingVoiceNotes.Add(new AnnouncementAttachment
                        {
                            FileName = fileName,
                            FilePath = $"/uploads/announcements/{fileName}",
                            FileSize = bytes.Length,
                            ContentType = string.IsNullOrEmpty(result.MimeType) ? "audio/webm" : result.MimeType,
                            UploadedAt = DateTime.UtcNow
                        });

                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Voice Recorded",
                            Detail = "Voice note recorded and attached successfully.",
                            Duration = 3000
                        });
                    }
                    else
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Warning,
                            Summary = "No Audio Received",
                            Detail = "No audio data was captured. Please speak into your microphone and try again.",
                            Duration = 4000
                        });
                    }

                    await InvokeAsync(StateHasChanged);
                }
            }
            catch (Exception ex)
            {
                IsRecordingVoice = false;
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Recording Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        protected void HandleFilesSelected(InputFileChangeEventArgs e)
        {
            foreach (var file in e.GetMultipleFiles(5))
            {
                PendingFiles.Add(file);
            }
        }

        protected void OpenCreateModal()
        {
            NewTitle = string.Empty;
            NewContent = string.Empty;
            NewDepartmentId = string.Empty;
            NewPriority = 0;
            NewIsPinned = false;
            PendingFiles.Clear();
            PendingVoiceNotes.Clear();
            IsRecordingVoice = false;
            ShowCreateModalState = true;
        }

        protected void CloseCreateModal()
        {
            ShowCreateModalState = false;
            IsRecordingVoice = false;
        }

        protected async Task CreateAnnouncementAsync()
        {
            if (string.IsNullOrWhiteSpace(NewTitle))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter an announcement title.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(NewContent))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter announcement details.",
                    Duration = 3000
                });
                return;
            }

            IsPublishing = true;
            try
            {
                var attachments = new List<AnnouncementAttachment>(PendingVoiceNotes);

                if (PendingFiles.Any())
                {
                    var uploadFolder = System.IO.Path.Combine(WebHostEnvironment.WebRootPath ?? "wwwroot", "uploads", "announcements");
                    System.IO.Directory.CreateDirectory(uploadFolder);

                    foreach (var file in PendingFiles)
                    {
                        var uniqueName = $"{Guid.NewGuid():N}_{file.Name}";
                        var fullPath = System.IO.Path.Combine(uploadFolder, uniqueName);

                        await using (var stream = new System.IO.FileStream(fullPath, System.IO.FileMode.Create))
                        {
                            await file.OpenReadStream(maxAllowedSize: 25 * 1024 * 1024).CopyToAsync(stream);
                        }

                        attachments.Add(new AnnouncementAttachment
                        {
                            FileName = file.Name,
                            FilePath = $"/uploads/announcements/{uniqueName}",
                            FileSize = file.Size,
                            ContentType = file.ContentType,
                            UploadedAt = DateTime.UtcNow
                        });
                    }
                }

                var ann = new Announcement
                {
                    Title = NewTitle.Trim(),
                    Content = NewContent.Trim(),
                    DepartmentId = string.IsNullOrWhiteSpace(NewDepartmentId) ? null : NewDepartmentId,
                    Priority = (AnnouncementPriority)NewPriority,
                    IsPinned = NewIsPinned,
                    CreatedBy = CurrentUserId,
                    Attachments = attachments
                };

                var id = await IntraOfficeService.CreateAnnouncementAsync(ann);
                if (id > 0)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Published",
                        Detail = "Announcement published successfully.",
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
                    Summary = "Publish Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsPublishing = false;
            }
        }

        protected async Task TogglePinAsync(int id)
        {
            var ann = AllAnnouncements.FirstOrDefault(a => a.Id == id);
            if (ann != null)
            {
                ann.IsPinned = !ann.IsPinned;
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Info,
                    Summary = ann.IsPinned ? "Pinned" : "Unpinned",
                    Detail = ann.IsPinned ? "Notice pinned to top." : "Notice unpinned.",
                    Duration = 2000
                });
            }
        }

        protected async Task AcknowledgeAsync(int announcementId)
        {
            try
            {
                var success = await IntraOfficeService.AcknowledgeAnnouncementAsync(announcementId, CurrentUserId);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Acknowledged",
                        Detail = "Thank you! Your acknowledgment has been recorded.",
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

        protected string GetCardBorderClass(AnnouncementPriority priority)
        {
            return priority switch
            {
                AnnouncementPriority.Critical => "border-danger",
                AnnouncementPriority.Important => "border-warning",
                _ => "border-light-subtle"
            };
        }

        protected string GetPriorityBadgeClass(AnnouncementPriority priority)
        {
            return priority switch
            {
                AnnouncementPriority.Critical => "bg-danger",
                AnnouncementPriority.Important => "bg-warning text-dark",
                _ => "bg-info"
            };
        }
    }
}
