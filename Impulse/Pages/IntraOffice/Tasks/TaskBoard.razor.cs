using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.IntraOffice.Tasks
{
    public partial class TaskBoard : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private IWhatsAppNotificationService WhatsAppService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;

        protected bool IsLoading { get; set; } = true;
        protected bool IsSaving { get; set; } = false;
        protected string ViewMode { get; set; } = "kanban";

        protected string CurrentUserId { get; set; } = string.Empty;
        protected string SearchTerm { get; set; } = string.Empty;
        protected string SelectedAssigneeFilter { get; set; } = string.Empty;
        protected string SelectedDepartmentFilter { get; set; } = string.Empty;

        protected List<TaskItem> AllTasksList { get; set; } = new();
        protected List<IntraUserProfile> UsersList { get; set; } = new();
        protected List<(string DeptId, string DeptName)> Departments { get; set; } = new();

        // Selected task for details modal
        protected TaskItem? SelectedTask { get; set; }
        protected string NewCommentText { get; set; } = string.Empty;

        // Create modal state & fields
        protected bool ShowCreateModalState { get; set; } = false;
        protected string NewTaskTitle { get; set; } = string.Empty;
        protected string NewTaskDescription { get; set; } = string.Empty;
        protected string NewTaskAssignee { get; set; } = string.Empty;
        protected string NewTaskDepartment { get; set; } = string.Empty;
        protected int NewTaskPriority { get; set; } = 1;
        protected DateTime? NewTaskDueDate { get; set; } = DateTime.Today.AddDays(2);
        protected bool NewTaskSendWhatsApp { get; set; } = true;

        protected List<TaskItem> FilteredTasks
        {
            get
            {
                var query = AllTasksList.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(SearchTerm))
                {
                    query = query.Where(t =>
                        t.Title.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        (t.Description?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        (t.AssigneeName?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false));
                }

                if (!string.IsNullOrWhiteSpace(SelectedAssigneeFilter))
                {
                    query = query.Where(t => t.AssignedTo == SelectedAssigneeFilter);
                }

                if (!string.IsNullOrWhiteSpace(SelectedDepartmentFilter))
                {
                    query = query.Where(t => t.DepartmentId == SelectedDepartmentFilter);
                }

                return query.ToList();
            }
        }

        protected List<TaskItem> GetTasksByStatus(TaskItemStatus status)
        {
            return FilteredTasks.Where(t => t.Status == status).ToList();
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
                AllTasksList = await IntraOfficeService.GetTasksAsync();
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

        protected async Task OpenTaskDetail(TaskItem task)
        {
            try
            {
                SelectedTask = await IntraOfficeService.GetTaskByIdAsync(task.Id) ?? task;
                NewCommentText = string.Empty;
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
                        Duration = 3000
                    });

                    if (SelectedTask != null && SelectedTask.Id == taskId)
                    {
                        SelectedTask.Status = newStatus;
                    }

                    var t = AllTasksList.FirstOrDefault(x => x.Id == taskId);
                    if (t != null)
                    {
                        t.Status = newStatus;
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Update Error",
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

            try
            {
                var comment = new TaskComment
                {
                    TaskId = SelectedTask.Id,
                    UserId = CurrentUserId,
                    Content = NewCommentText.Trim()
                };

                var commentId = await IntraOfficeService.AddTaskCommentAsync(comment);
                if (commentId > 0)
                {
                    comment.Id = commentId;
                    comment.UserName = CurrentUserId;
                    comment.CreatedAt = DateTime.UtcNow;
                    SelectedTask.Comments.Add(comment);
                    NewCommentText = string.Empty;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Comment Error",
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
            NewTaskSendWhatsApp = true;
            ShowCreateModalState = true;
        }

        protected void CloseCreateModal()
        {
            ShowCreateModalState = false;
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
                var task = new TaskItem
                {
                    Title = NewTaskTitle.Trim(),
                    Description = string.IsNullOrWhiteSpace(NewTaskDescription) ? null : NewTaskDescription.Trim(),
                    AssignedTo = string.IsNullOrWhiteSpace(NewTaskAssignee) ? null : NewTaskAssignee,
                    AssignedBy = CurrentUserId,
                    DepartmentId = string.IsNullOrWhiteSpace(NewTaskDepartment) ? null : NewTaskDepartment,
                    Priority = (TaskPriority)NewTaskPriority,
                    Status = TaskItemStatus.Pending,
                    DueDate = NewTaskDueDate
                };

                var taskId = await IntraOfficeService.CreateTaskAsync(task);
                if (taskId > 0)
                {
                    // Optionally send WhatsApp notification to assignee
                    if (NewTaskSendWhatsApp && !string.IsNullOrEmpty(task.AssignedTo))
                    {
                        var assignee = UsersList.FirstOrDefault(u => u.UserName == task.AssignedTo);
                        if (assignee != null && !string.IsNullOrWhiteSpace(assignee.CellNo))
                        {
                            _ = WhatsAppService.SendTaskNotificationAsync(
                                assignee.CellNo,
                                task.Title,
                                assignee.EmployeeName ?? assignee.FullUserName ?? assignee.UserName,
                                task.Priority.ToString(),
                                task.Description
                            );
                        }
                    }

                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Created",
                        Detail = "Task created successfully.",
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
                    Summary = "Save Error",
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
                TaskItemStatus.InProgress => "bg-primary",
                TaskItemStatus.Cancelled => "bg-danger",
                _ => "bg-warning text-dark"
            };
        }
    }
}
