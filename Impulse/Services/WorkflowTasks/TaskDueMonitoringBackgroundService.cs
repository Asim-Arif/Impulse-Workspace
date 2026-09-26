using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.IntraOffice;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.WorkflowTasks
{
    public class TaskDueMonitoringBackgroundService : BackgroundService
    {
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly ILogger<TaskDueMonitoringBackgroundService> _logger;
        private static readonly TimeSpan CheckInterval = TimeSpan.FromMinutes(2);

        public TaskDueMonitoringBackgroundService(
            IServiceScopeFactory scopeFactory,
            ILogger<TaskDueMonitoringBackgroundService> logger)
        {
            _scopeFactory = scopeFactory;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("TaskDueMonitoringBackgroundService has started.");

            // Wait 15 seconds after app startup before first run
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(15), stoppingToken);
            }
            catch (OperationCanceledException)
            {
                return;
            }

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    using var scope = _scopeFactory.CreateScope();
                    var dataAccess = scope.ServiceProvider.GetRequiredService<IIntraOfficeDataAccess>();
                    var notificationService = scope.ServiceProvider.GetRequiredService<IAppNotificationService>();

                    await Process70PercentDueWarningsAsync(dataAccess, notificationService, stoppingToken);
                    await ProcessOverdueAlertsAsync(dataAccess, notificationService, stoppingToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Unexpected error in TaskDueMonitoringBackgroundService execution cycle.");
                }

                try
                {
                    await Task.Delay(CheckInterval, stoppingToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }

            _logger.LogInformation("TaskDueMonitoringBackgroundService is stopping.");
        }

        private async Task Process70PercentDueWarningsAsync(
            IIntraOfficeDataAccess dataAccess,
            IAppNotificationService notificationService,
            CancellationToken ct)
        {
            var pendingTasks = await dataAccess.GetTasksPending70PercentDueWarningAsync();
            if (!pendingTasks.Any()) return;

            _logger.LogInformation("Found {Count} task(s) meeting 70% due threshold.", pendingTasks.Count);

            foreach (var task in pendingTasks)
            {
                if (ct.IsCancellationRequested) break;

                try
                {
                    var targetUsers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (task.AssigneeUserNames.Any())
                    {
                        foreach (var u in task.AssigneeUserNames)
                        {
                            if (!string.IsNullOrWhiteSpace(u)) targetUsers.Add(u.Trim());
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(task.AssignedTo))
                    {
                        foreach (var part in task.AssignedTo.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            if (!string.IsNullOrWhiteSpace(part)) targetUsers.Add(part.Trim());
                        }
                    }

                    if (!targetUsers.Any() && !string.IsNullOrWhiteSpace(task.AssignedBy))
                    {
                        targetUsers.Add(task.AssignedBy.Trim());
                    }

                    string title = $"⚠️ Task Due Soon (70% Elapsed): {task.Title}";
                    string message = $"Task '{task.Title}' has used 70% of its allotted time. Due date is {task.DueDate:dd-MMM-yyyy HH:mm}. Please expedite completion.";
                    string actionUrl = !string.IsNullOrEmpty(task.ActionUrl) ? task.ActionUrl : "/office/tasks";

                    foreach (var userName in targetUsers)
                    {
                        await notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = title,
                            Message = message,
                            SenderName = "System",
                            TargetUserId = userName,
                            ActionUrl = actionUrl
                        });
                    }

                    // Append notice into TaskComments
                    await dataAccess.AddTaskCommentAsync(new TaskComment
                    {
                        TaskId = task.Id,
                        UserId = "System",
                        Content = $"Automated Notice: 70% of allotted timeline has elapsed. Deadline is {task.DueDate:dd-MMM-yyyy HH:mm}.",
                        CreatedAt = DateTime.UtcNow
                    });

                    await dataAccess.MarkDueWarningSentAsync(task.Id);
                    _logger.LogInformation("Dispatched 70% due warning for Task #{TaskId} to {UserCount} user(s).", task.Id, targetUsers.Count);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing 70% due warning for Task #{TaskId}", task.Id);
                }
            }
        }

        private async Task ProcessOverdueAlertsAsync(
            IIntraOfficeDataAccess dataAccess,
            IAppNotificationService notificationService,
            CancellationToken ct)
        {
            var overdueTasks = await dataAccess.GetTasksPendingOverdueAlertAsync();
            if (!overdueTasks.Any()) return;

            _logger.LogInformation("Found {Count} overdue task(s) requiring supervisor escalation.", overdueTasks.Count);

            foreach (var task in overdueTasks)
            {
                if (ct.IsCancellationRequested) break;

                try
                {
                    var supervisors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    // If linked to a lot hub task (SourceEntityType = 'LotHub' and RefId = 'LotNo:HubName')
                    if (string.Equals(task.SourceEntityType, "LotHub", StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrWhiteSpace(task.SourceEntityRefId)
                        && task.SourceEntityRefId.Contains(':'))
                    {
                        var parts = task.SourceEntityRefId.Split(':');
                        string lotNo = parts[0];
                        string hubName = parts[1];

                        var hubSupervisors = await dataAccess.GetHubSupervisorsForLotTaskAsync(lotNo, hubName);
                        foreach (var s in hubSupervisors)
                        {
                            if (!string.IsNullOrWhiteSpace(s)) supervisors.Add(s.Trim());
                        }
                    }

                    // Fallback to task assignees / assigner / administrator if no specific hub supervisor found
                    if (!supervisors.Any())
                    {
                        foreach (var u in task.AssigneeUserNames)
                        {
                            if (!string.IsNullOrWhiteSpace(u)) supervisors.Add(u.Trim());
                        }

                        if (!string.IsNullOrWhiteSpace(task.AssignedTo))
                        {
                            foreach (var part in task.AssignedTo.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                if (!string.IsNullOrWhiteSpace(part)) supervisors.Add(part.Trim());
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(task.AssignedBy))
                        {
                            supervisors.Add(task.AssignedBy.Trim());
                        }

                        supervisors.Add("Administrator");
                    }

                    var overdueDuration = task.DueDate.HasValue && DateTime.UtcNow > task.DueDate.Value
                        ? DateTime.UtcNow - task.DueDate.Value
                        : TimeSpan.Zero;
                    string formattedOverdue = FormatTimeSpan(overdueDuration);

                    string overdueTitle = $"🚨 OVERDUE Task Alert: {task.Title}";
                    string overdueMessage = $"Task '{task.Title}' was not completed by its deadline ({task.DueDate:dd-MMM-yyyy HH:mm}) and is now OVERDUE by {formattedOverdue}.";
                    string actionUrl = !string.IsNullOrEmpty(task.ActionUrl) ? task.ActionUrl : "/office/tasks";

                    // 1. Send high-priority Overdue Notification to supervisors
                    foreach (var supervisor in supervisors)
                    {
                        await notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = overdueTitle,
                            Message = overdueMessage,
                            SenderName = "System",
                            TargetUserId = supervisor,
                            ActionUrl = actionUrl
                        });
                    }

                    // 2. Append comment to original task
                    string supervisorList = string.Join(", ", supervisors);
                    await dataAccess.AddTaskCommentAsync(new TaskComment
                    {
                        TaskId = task.Id,
                        UserId = "System",
                        Content = $"Automated Alert: Task is OVERDUE by {formattedOverdue}. Overdue alert dispatched to Hub Supervisor(s): {supervisorList}.",
                        CreatedAt = DateTime.UtcNow
                    });

                    // 3. Mark OverdueWarningSent = 1
                    await dataAccess.MarkOverdueWarningSentAsync(task.Id);
                    _logger.LogInformation("Dispatched Overdue alert for Task #{TaskId} to supervisor(s): {Supervisors}.", task.Id, supervisorList);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing overdue alert for Task #{TaskId}", task.Id);
                }
            }
        }

        private static string FormatTimeSpan(TimeSpan ts)
        {
            if (ts.TotalDays >= 1)
            {
                return $"{ts.Days}d {ts.Hours}h {ts.Minutes}m";
            }
            if (ts.TotalHours >= 1)
            {
                return $"{ts.Hours}h {ts.Minutes}m";
            }
            return $"{Math.Max(1, ts.Minutes)}m";
        }
    }
}
