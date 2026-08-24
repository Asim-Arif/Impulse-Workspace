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
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Hosting;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.IntraOffice.Tasks
{
    public partial class TaskBoard : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private IWhatsAppNotificationService WhatsAppService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private IWebHostEnvironment WebHostEnvironment { get; set; } = null!;
        [Inject] private IJSRuntime JS { get; set; } = null!;

        protected string CurrentUserId { get; set; } = string.Empty;
        protected bool IsLoading { get; set; } = true;
        protected bool IsSaving { get; set; } = false;
        protected bool IsRecordingVoice { get; set; } = false;
        protected string ViewMode { get; set; } = "kanban";

        protected List<TaskItem> AllTasks { get; set; } = new();
        protected List<IntraUserProfile> UsersList { get; set; } = new();
        protected List<(string DeptId, string DeptName)> Departments { get; set; } = new();

        // Filters
        protected string SearchTerm { get; set; } = string.Empty;
        protected string SelectedAssigneeFilter { get; set; } = string.Empty;
        protected int SelectedPriorityFilter { get; set; } = -1;
        protected string SelectedDepartmentFilter { get; set; } = string.Empty;

        // Detail Modal
        protected TaskItem? SelectedTask { get; set; }
        protected string NewCommentText { get; set; } = string.Empty;

        // Create Modal
        protected bool ShowCreateModalState { get; set; } = false;
        protected string NewTaskTitle { get; set; } = string.Empty;
        protected string NewTaskDescription { get; set; } = string.Empty;
        protected string NewTaskAssignee { get; set; } = string.Empty;
        protected string NewTaskDepartment { get; set; } = string.Empty;
        protected int NewTaskPriority { get; set; } = 1;
        protected DateTime? NewTaskDueDate { get; set; }
        protected bool NewTaskSendWhatsApp { get; set; } = false;

        protected List<IBrowserFile> PendingFiles { get; set; } = new();
        protected List<TaskAttachment> PendingVoiceNotes { get; set; } = new();

        public class AudioRecordingResult
        {
            public string Base64 { get; set; } = string.Empty;
            public string MimeType { get; set; } = string.Empty;
        }

        protected List<TaskItem> FilteredTasks
        {
            get
            {
                var query = AllTasks.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(SelectedAssigneeFilter))
                {
                    query = query.Where(t => string.Equals(t.AssignedTo, SelectedAssigneeFilter, StringComparison.OrdinalIgnoreCase));
                }

                if (!string.IsNullOrWhiteSpace(SelectedDepartmentFilter))
                {
                    query = query.Where(t => string.Equals(t.DepartmentId, SelectedDepartmentFilter, StringComparison.OrdinalIgnoreCase));
                }

                if (SelectedPriorityFilter >= 0)
                {
                    query = query.Where(t => (int)t.Priority == SelectedPriorityFilter);
                }

                if (!string.IsNullOrWhiteSpace(SearchTerm))
                {
                    query = query.Where(t =>
                        t.Title.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        (t.Description?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        (t.AssigneeName?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false));
                }

                return query.ToList();
            }
        }

        protected List<TaskItem> GetTasksByStatus(TaskItemStatus status)
        {
            return FilteredTasks.Where(t => t.Status == status).OrderByDescending(t => t.Priority).ThenBy(t => t.DueDate).ToList();
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
                UsersList = await IntraOfficeService.GetActiveUsersAsync();
                Departments = await IntraOfficeService.GetDepartmentsAsync();
                AllTasks = await IntraOfficeService.GetTasksAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Tasks",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected async Task OpenTaskDetail(TaskItem task)
        {
            SelectedTask = await IntraOfficeService.GetTaskByIdAsync(task.Id) ?? task;
            NewCommentText = string.Empty;
        }

        protected void CloseTaskDetail()
        {
            SelectedTask = null;
        }

        protected async Task ChangeStatusAsync(int taskId, TaskItemStatus newStatus)
        {
            try
            {
                var success = await IntraOfficeService.UpdateTaskStatusAsync(taskId, newStatus);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Status Updated",
                        Detail = $"Task moved to {newStatus}.",
                        Duration = 2500
                    });

                    if (SelectedTask != null && SelectedTask.Id == taskId)
                    {
                        SelectedTask.Status = newStatus;
                    }

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

        protected async Task HandleCommentKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter" && !string.IsNullOrWhiteSpace(NewCommentText))
            {
                await AddCommentAsync();
            }
        }

        protected async Task AddCommentAsync()
        {
            if (SelectedTask == null || string.IsNullOrWhiteSpace(NewCommentText)) return;

            var text = NewCommentText.Trim();
            NewCommentText = string.Empty;

            try
            {
                var comment = new TaskComment
                {
                    TaskId = SelectedTask.Id,
                    UserId = CurrentUserId,
                    Content = text,
                    CreatedAt = DateTime.UtcNow
                };

                var id = await IntraOfficeService.AddTaskCommentAsync(comment);
                if (id > 0)
                {
                    comment.Id = id;
                    comment.UserName = CurrentUserId;
                    SelectedTask.Comments.Add(comment);
                    StateHasChanged();
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
                        var uploadFolder = System.IO.Path.Combine(WebHostEnvironment.WebRootPath ?? "wwwroot", "uploads", "tasks");
                        System.IO.Directory.CreateDirectory(uploadFolder);

                        var fileName = $"TaskVoice_{DateTime.Now:yyyyMMdd_HHmmss}.webm";
                        var filePath = System.IO.Path.Combine(uploadFolder, fileName);
                        var bytes = Convert.FromBase64String(result.Base64);

                        await System.IO.File.WriteAllBytesAsync(filePath, bytes);

                        PendingVoiceNotes.Add(new TaskAttachment
                        {
                            FileName = fileName,
                            FilePath = $"/uploads/tasks/{fileName}",
                            FileSize = bytes.Length,
                            ContentType = string.IsNullOrEmpty(result.MimeType) ? "audio/webm" : result.MimeType,
                            UploadedAt = DateTime.UtcNow
                        });

                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Voice Recorded",
                            Detail = "Voice note attached to task.",
                            Duration = 3000
                        });
                    }
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

        protected void HandleNewAttachments(InputFileChangeEventArgs e)
        {
            foreach (var file in e.GetMultipleFiles(5))
            {
                PendingFiles.Add(file);
            }
        }

        protected async Task HandleDetailFileUpload(InputFileChangeEventArgs e)
        {
            if (SelectedTask == null) return;

            try
            {
                var uploadFolder = System.IO.Path.Combine(WebHostEnvironment.WebRootPath ?? "wwwroot", "uploads", "tasks");
                System.IO.Directory.CreateDirectory(uploadFolder);

                foreach (var file in e.GetMultipleFiles(5))
                {
                    var uniqueName = $"{Guid.NewGuid():N}_{file.Name}";
                    var fullPath = System.IO.Path.Combine(uploadFolder, uniqueName);

                    await using (var stream = new System.IO.FileStream(fullPath, System.IO.FileMode.Create))
                    {
                        await file.OpenReadStream(maxAllowedSize: 25 * 1024 * 1024).CopyToAsync(stream);
                    }

                    var att = new TaskAttachment
                    {
                        TaskId = SelectedTask.Id,
                        FileName = file.Name,
                        FilePath = $"/uploads/tasks/{uniqueName}",
                        FileSize = file.Size,
                        ContentType = file.ContentType,
                        UploadedAt = DateTime.UtcNow
                    };

                    SelectedTask.Attachments.Add(att);
                }

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Attached",
                    Detail = "Document attached to task.",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Upload Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        protected void OpenCreateModal()
        {
            NewTaskTitle = string.Empty;
            NewTaskDescription = string.Empty;
            NewTaskAssignee = string.Empty;
            NewTaskDepartment = string.Empty;
            NewTaskPriority = 1;
            NewTaskDueDate = DateTime.Today.AddDays(2);
            NewTaskSendWhatsApp = false;
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

        protected async Task SaveTaskAsync()
        {
            if (string.IsNullOrWhiteSpace(NewTaskTitle))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter a task title.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            try
            {
                var attachments = new List<TaskAttachment>(PendingVoiceNotes);

                if (PendingFiles.Any())
                {
                    var uploadFolder = System.IO.Path.Combine(WebHostEnvironment.WebRootPath ?? "wwwroot", "uploads", "tasks");
                    System.IO.Directory.CreateDirectory(uploadFolder);

                    foreach (var file in PendingFiles)
                    {
                        var uniqueName = $"{Guid.NewGuid():N}_{file.Name}";
                        var fullPath = System.IO.Path.Combine(uploadFolder, uniqueName);

                        await using (var stream = new System.IO.FileStream(fullPath, System.IO.FileMode.Create))
                        {
                            await file.OpenReadStream(maxAllowedSize: 25 * 1024 * 1024).CopyToAsync(stream);
                        }

                        attachments.Add(new TaskAttachment
                        {
                            FileName = file.Name,
                            FilePath = $"/uploads/tasks/{uniqueName}",
                            FileSize = file.Size,
                            ContentType = file.ContentType,
                            UploadedAt = DateTime.UtcNow
                        });
                    }
                }

                var task = new TaskItem
                {
                    Title = NewTaskTitle.Trim(),
                    Description = string.IsNullOrWhiteSpace(NewTaskDescription) ? null : NewTaskDescription.Trim(),
                    AssignedTo = string.IsNullOrWhiteSpace(NewTaskAssignee) ? null : NewTaskAssignee,
                    DepartmentId = string.IsNullOrWhiteSpace(NewTaskDepartment) ? null : NewTaskDepartment,
                    Priority = (TaskPriority)NewTaskPriority,
                    Status = TaskItemStatus.Pending,
                    DueDate = NewTaskDueDate,
                    AssignedBy = CurrentUserId,
                    CreatedAt = DateTime.UtcNow,
                    Attachments = attachments
                };

                var id = await IntraOfficeService.CreateTaskAsync(task);
                if (id > 0)
                {
                    task.Id = id;

                    // Send WhatsApp if checked
                    if (NewTaskSendWhatsApp && !string.IsNullOrEmpty(task.AssignedTo))
                    {
                        var assignee = UsersList.FirstOrDefault(u => u.UserName == task.AssignedTo);
                        if (!string.IsNullOrEmpty(assignee?.CellNo))
                        {
                            var msg = $"*New Task Assigned: {task.Title}*\nPriority: {task.Priority}\nDue: {task.DueDate:MMM dd, yyyy}\nBy: {CurrentUserId}";
                            _ = WhatsAppService.SendNotificationAsync(assignee.CellNo, msg);
                        }
                    }

                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Task Created",
                        Detail = "Task was successfully created and assigned.",
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
            finally
            {
                IsSaving = false;
            }
        }

        protected string GetPriorityBadgeClass(TaskPriority priority)
        {
            return priority switch
            {
                TaskPriority.Urgent => "bg-danger",
                TaskPriority.High => "bg-warning text-dark",
                TaskPriority.Normal => "bg-primary",
                _ => "bg-secondary"
            };
        }

        protected string GetStatusBadgeClass(TaskItemStatus status)
        {
            return status switch
            {
                TaskItemStatus.Completed => "bg-success",
                TaskItemStatus.InProgress => "bg-info",
                TaskItemStatus.Cancelled => "bg-danger",
                _ => "bg-secondary"
            };
        }
    }
}
