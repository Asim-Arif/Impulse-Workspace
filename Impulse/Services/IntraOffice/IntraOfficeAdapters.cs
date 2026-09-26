using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Microsoft.AspNetCore.Components.Forms;

namespace Impulse.Services.IntraOffice
{
    // -------------------------------------------------------------
    // Channel Service
    // -------------------------------------------------------------
    public interface IChannelService
    {
        Task<List<Channel>> GetAllChannelsAsync();
        Task<List<Channel>> GetUserChannelsAsync(string userId);
        Task<Channel?> GetChannelByIdAsync(int id);
        Task<Channel> CreateChannelAsync(Channel channel);
        Task UpdateChannelAsync(Channel channel);
        Task DeleteChannelAsync(int id);
        Task AddMemberAsync(int channelId, string userId, bool isAdmin = false);
        Task RemoveMemberAsync(int channelId, string userId);
        Task<List<ChannelMember>> GetChannelMembersAsync(int channelId);
        Task<bool> IsMemberAsync(int channelId, string userId);
    }

    public class ChannelService : IChannelService
    {
        private readonly IIntraOfficeService _intra;
        public ChannelService(IIntraOfficeService intra) => _intra = intra;

        public async Task<List<Channel>> GetAllChannelsAsync() => await _intra.GetChannelsForUserAsync("");
        public async Task<List<Channel>> GetUserChannelsAsync(string userId) => await _intra.GetChannelsForUserAsync(userId);
        public async Task<Channel?> GetChannelByIdAsync(int id) => await _intra.GetChannelByIdAsync(id);
        public async Task<Channel> CreateChannelAsync(Channel channel)
        {
            var id = await _intra.CreateChannelAsync(channel);
            channel.Id = id;
            return channel;
        }
        public Task UpdateChannelAsync(Channel channel) => Task.CompletedTask;
        public Task DeleteChannelAsync(int id) => Task.CompletedTask;
        public async Task AddMemberAsync(int channelId, string userId, bool isAdmin = false) => await _intra.AddMemberToChannelAsync(channelId, userId, isAdmin);
        public async Task RemoveMemberAsync(int channelId, string userId) => await _intra.RemoveMemberFromChannelAsync(channelId, userId);
        public Task<List<ChannelMember>> GetChannelMembersAsync(int channelId) => Task.FromResult(new List<ChannelMember>());
        public Task<bool> IsMemberAsync(int channelId, string userId) => Task.FromResult(true);
    }

    // -------------------------------------------------------------
    // Message Service
    // -------------------------------------------------------------
    public interface IMessageService
    {
        Task<List<Message>> GetChannelMessagesAsync(int channelId, int page = 1, int pageSize = 50);
        Task<List<Message>> GetDirectMessagesAsync(string userId, string otherUserId, int page = 1, int pageSize = 50);
        Task<Message> SendMessageAsync(Message message);
        Task<Message?> GetMessageByIdAsync(long id);
        Task MarkAsReadAsync(long messageId);
        Task MarkAllAsReadAsync(string userId, int? channelId = null, string? senderId = null);
        Task<int> GetUnreadCountAsync(string userId);
        Task<Dictionary<string, int>> GetUnreadDirectMessageCountsAsync(string userId);
        Task DeleteMessageAsync(long id);
        Task<List<Message>> GetRecentConversationsAsync(string userId);
    }

    public class MessageService : IMessageService
    {
        private readonly IIntraOfficeService _intra;
        private readonly IAppNotificationService _notifications;

        public MessageService(IIntraOfficeService intra, IAppNotificationService notifications)
        {
            _intra = intra;
            _notifications = notifications;
        }

        public async Task<List<Message>> GetChannelMessagesAsync(int channelId, int page = 1, int pageSize = 50)
            => await _intra.GetChannelMessagesAsync(channelId, pageSize);

        public async Task<List<Message>> GetDirectMessagesAsync(string userId, string otherUserId, int page = 1, int pageSize = 50)
            => await _intra.GetDirectMessagesAsync(userId, otherUserId, pageSize);

        public async Task<Message> SendMessageAsync(Message message)
        {
            var id = await _intra.SendMessageAsync(message);
            message.Id = id;

            var senderName = message.SenderName ?? message.SenderId;
            if (string.IsNullOrEmpty(message.SenderName) && !string.IsNullOrEmpty(message.SenderId))
            {
                try
                {
                    var senderProfile = await _intra.GetUserProfileAsync(message.SenderId);
                    if (senderProfile != null) senderName = senderProfile.FullName;
                }
                catch { }
            }

            var preview = !string.IsNullOrWhiteSpace(message.Content)
                ? (message.Content.Length > 80 ? message.Content.Substring(0, 80) + "..." : message.Content)
                : (message.Attachments != null && message.Attachments.Any() ? "[Attachment]" : "Sent a message");

            if (!string.IsNullOrEmpty(message.ReceiverId))
            {
                _ = _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Message,
                    Title = $"New Message from {senderName}",
                    Message = preview,
                    SenderName = senderName,
                    TargetUserId = message.ReceiverId,
                    ActionUrl = $"/office/messages/{message.SenderId}"
                });
            }
            else if (message.ChannelId.HasValue && message.ChannelId > 0)
            {
                var channelName = "Discussion";
                try
                {
                    var channel = await _intra.GetChannelByIdAsync(message.ChannelId.Value);
                    if (channel != null && !string.IsNullOrEmpty(channel.Name)) channelName = channel.Name;
                }
                catch { }

                _ = _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Message,
                    Title = $"#{channelName} • {senderName}",
                    Message = preview,
                    SenderName = senderName,
                    TargetUserId = null, // broadcast to all channel participants
                    ActionUrl = $"/office/chat/{message.ChannelId.Value}"
                });
            }
            return message;
        }

        public Task<Message?> GetMessageByIdAsync(long id) => Task.FromResult<Message?>(null);
        public Task MarkAsReadAsync(long messageId) => Task.CompletedTask;
        public async Task MarkAllAsReadAsync(string userId, int? channelId = null, string? senderId = null)
        {
            await _intra.MarkMessagesAsReadAsync(userId, senderId, channelId);

            if (!string.IsNullOrEmpty(senderId) && !string.Equals(senderId, userId, StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var reader = await _intra.GetUserProfileAsync(userId);
                    var readerName = reader?.FullName ?? "Someone";
                    _ = _notifications.SendNotificationAsync(new AppNotification
                    {
                        Category = NotificationCategory.Message,
                        Title = $"✓✓ {readerName} read your messages",
                        Message = $"{readerName} has seen your direct messages.",
                        SenderName = readerName,
                        TargetUserId = senderId,
                        ActionUrl = $"/office/messages/{userId}",
                        IsReadReceipt = true
                    });
                }
                catch { }
            }
        }

        public async Task<int> GetUnreadCountAsync(string userId)
        {
            var summaries = await _intra.GetDirectChatSummariesAsync(userId);
            return summaries.Sum(s => s.UnreadCount);
        }

        public async Task<Dictionary<string, int>> GetUnreadDirectMessageCountsAsync(string userId)
        {
            var summaries = await _intra.GetDirectChatSummariesAsync(userId);
            return summaries.ToDictionary(s => s.OtherUserId, s => s.UnreadCount);
        }

        public async Task DeleteMessageAsync(long id) => await _intra.DeleteMessageAsync(id, "");
        public async Task<List<Message>> GetRecentConversationsAsync(string userId)
        {
            var summaries = await _intra.GetDirectChatSummariesAsync(userId);
            return summaries.Select(s => new Message
            {
                SenderId = s.OtherUserId,
                ReceiverId = userId,
                Content = s.LastMessage ?? "",
                SentAt = s.LastMessageAt
            }).ToList();
        }
    }

    // -------------------------------------------------------------
    // Announcement Service
    // -------------------------------------------------------------
    public interface IAnnouncementService
    {
        event Action? OnAnnouncementUpdated;
        Task<List<Announcement>> GetAllAnnouncementsAsync(int? departmentId = null);
        Task<List<Announcement>> GetActiveAnnouncementsAsync(int? departmentId = null);
        Task<Announcement?> GetAnnouncementByIdAsync(int id);
        Task<Announcement> CreateAnnouncementAsync(Announcement announcement);
        Task UpdateAnnouncementAsync(Announcement announcement);
        Task DeleteAnnouncementAsync(int id);
        Task TogglePinAsync(int id);
        Task<int> GetUnreadAnnouncementCountAsync(string userId);
        Task MarkAnnouncementsAsReadAsync(string userId);
        Task AcknowledgeAnnouncementAsync(int announcementId, string userId);
        Task<List<int>> GetUserAcknowledgedAnnouncementIdsAsync(string userId);
    }

    public class AnnouncementService : IAnnouncementService
    {
        private static event Action? _onAnnouncementUpdated;
        public event Action? OnAnnouncementUpdated
        {
            add => _onAnnouncementUpdated += value;
            remove => _onAnnouncementUpdated -= value;
        }

        private readonly IIntraOfficeService _intra;
        private readonly IAppNotificationService _notifications;

        public AnnouncementService(IIntraOfficeService intra, IAppNotificationService notifications)
        {
            _intra = intra;
            _notifications = notifications;
        }

        public async Task<List<Announcement>> GetAllAnnouncementsAsync(int? departmentId = null)
            => await _intra.GetAnnouncementsAsync("", departmentId?.ToString());

        public async Task<List<Announcement>> GetActiveAnnouncementsAsync(int? departmentId = null)
            => await _intra.GetAnnouncementsAsync("", departmentId?.ToString());

        public async Task<Announcement?> GetAnnouncementByIdAsync(int id) => await _intra.GetAnnouncementByIdAsync(id, "");

        public async Task<Announcement> CreateAnnouncementAsync(Announcement announcement)
        {
            var id = await _intra.CreateAnnouncementAsync(announcement);
            announcement.Id = id;

            var authorProfile = !string.IsNullOrEmpty(announcement.CreatedBy) ? await _intra.GetUserProfileAsync(announcement.CreatedBy) : null;
            var authorName = authorProfile?.FullName ?? "Management";
            var preview = !string.IsNullOrWhiteSpace(announcement.Content)
                ? (announcement.Content.Length > 100 ? announcement.Content.Substring(0, 100) + "..." : announcement.Content)
                : announcement.Title;

            _ = _notifications.SendNotificationAsync(new AppNotification
            {
                Category = NotificationCategory.Announcement,
                Title = $"📢 Announcement: {announcement.Title}",
                Message = preview,
                SenderName = authorName,
                TargetUserId = null, // broadcast to all
                ActionUrl = "/office/announcements"
            });
            return announcement;
        }

        public Task UpdateAnnouncementAsync(Announcement announcement) => Task.CompletedTask;
        public Task DeleteAnnouncementAsync(int id) => Task.CompletedTask;
        public Task TogglePinAsync(int id) => Task.CompletedTask;
        public async Task<int> GetUnreadAnnouncementCountAsync(string userId)
        {
            var all = await _intra.GetAnnouncementsAsync(userId);
            return all.Count(a => !a.IsAcknowledgedByCurrentUser);
        }
        public Task MarkAnnouncementsAsReadAsync(string userId) => Task.CompletedTask;
        public async Task AcknowledgeAnnouncementAsync(int announcementId, string userId)
        {
            await _intra.AcknowledgeAnnouncementAsync(announcementId, userId);

            try
            {
                var announcement = await _intra.GetAnnouncementByIdAsync(announcementId, userId);
                if (announcement != null && !string.IsNullOrEmpty(announcement.CreatedBy) && !string.Equals(announcement.CreatedBy, userId, StringComparison.OrdinalIgnoreCase))
                {
                    var reader = await _intra.GetUserProfileAsync(userId);
                    var readerName = reader?.FullName ?? "Someone";
                    _ = _notifications.SendNotificationAsync(new AppNotification
                    {
                        Category = NotificationCategory.Announcement,
                        Title = $"✓ {readerName} acknowledged your announcement",
                        Message = $"\"{announcement.Title}\" was acknowledged by {readerName}.",
                        SenderName = readerName,
                        TargetUserId = announcement.CreatedBy,
                        ActionUrl = "/office/announcements",
                        IsReadReceipt = true
                    });
                }
            }
            catch { }
        }

        public Task<List<int>> GetUserAcknowledgedAnnouncementIdsAsync(string userId)
            => Task.FromResult(new List<int>());
    }

    // -------------------------------------------------------------
    // Task Service
    // -------------------------------------------------------------
    public interface ITaskService
    {
        Task<List<TaskItem>> GetAllTasksAsync(int? departmentId = null);
        Task<List<TaskItem>> GetUserTasksAsync(string userId);
        Task<List<TaskItem>> GetTasksByStatusAsync(TaskItemStatus status, int? departmentId = null);
        Task<TaskItem?> GetTaskByIdAsync(int id);
        Task<TaskItem> CreateTaskAsync(TaskItem task, bool sendWhatsApp = true, IEnumerable<IBrowserFile>? attachments = null, bool sendEmail = true, List<string>? allAssigneeIds = null);
        Task UpdateTaskAsync(TaskItem task);
        Task UpdateStatusAsync(int taskId, TaskItemStatus status, string? changedByUserId = null);
        Task DeleteTaskAsync(int id);
        Task<TaskComment> AddCommentAsync(TaskComment comment);
        Task<List<TaskComment>> GetTaskCommentsAsync(int taskId);
        Task<TaskAttachment> AddAttachmentAsync(TaskAttachment attachment);
        Task<int> GetUnreadTaskCountAsync(string userId);
        Task MarkTaskAsReadAsync(int taskId, string userId);
        Task MarkAllTasksAsReadAsync(string userId);
    }

    public class TaskService : ITaskService
    {
        private readonly IIntraOfficeService _intra;
        private readonly IAppNotificationService _notifications;
        private readonly IWhatsAppService _whatsApp;

        public TaskService(IIntraOfficeService intra, IAppNotificationService notifications, IWhatsAppService whatsApp)
        {
            _intra = intra;
            _notifications = notifications;
            _whatsApp = whatsApp;
        }

        public async Task<List<TaskItem>> GetAllTasksAsync(int? departmentId = null) => await _intra.GetTasksAsync(departmentId: departmentId?.ToString());
        public async Task<List<TaskItem>> GetUserTasksAsync(string userId) => await _intra.GetTasksAsync(assignedTo: userId);
        public async Task<List<TaskItem>> GetTasksByStatusAsync(TaskItemStatus status, int? departmentId = null) => await _intra.GetTasksAsync(departmentId: departmentId?.ToString(), status: status);
        public async Task<TaskItem?> GetTaskByIdAsync(int id) => await _intra.GetTaskByIdAsync(id);

        public async Task<TaskItem> CreateTaskAsync(TaskItem task, bool sendWhatsApp = true, IEnumerable<IBrowserFile>? attachments = null, bool sendEmail = true, List<string>? allAssigneeIds = null)
        {
            if (string.IsNullOrEmpty(task.AdditionalAssigneeIds) && allAssigneeIds != null && allAssigneeIds.Count > 1)
            {
                var additional = allAssigneeIds.Where(a => !string.Equals(a, task.AssignedTo, StringComparison.OrdinalIgnoreCase)).ToList();
                if (additional.Any())
                {
                    task.AdditionalAssigneeIds = string.Join(",", additional);
                }
            }

            var id = await _intra.CreateTaskAsync(task);
            task.Id = id;

            var assignerProfile = !string.IsNullOrEmpty(task.AssignedBy) ? await _intra.GetUserProfileAsync(task.AssignedBy) : null;
            var assignerName = assignerProfile?.FullName ?? task.AssignedBy ?? "Management";

            var targetAssignees = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(task.AssignedTo)) targetAssignees.Add(task.AssignedTo);
            if (allAssigneeIds != null)
            {
                foreach (var a in allAssigneeIds)
                {
                    if (!string.IsNullOrEmpty(a)) targetAssignees.Add(a);
                }
            }

            foreach (var assigneeId in targetAssignees)
            {
                _ = _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Task,
                    Title = $"Task Assigned: {task.Title}",
                    Message = $"Priority: {task.Priority} • Assigned by {assignerName}",
                    SenderName = assignerName,
                    TargetUserId = assigneeId,
                    ActionUrl = "/office/tasks"
                });

                if (sendWhatsApp)
                {
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            var userProfile = await _intra.GetUserProfileAsync(assigneeId);
                            if (!string.IsNullOrEmpty(userProfile?.CellNo))
                            {
                                await _whatsApp.SendTaskNotificationAsync(userProfile.CellNo, task.Title, userProfile.FullName ?? assigneeId, task.Priority.ToString(), null, task.Description);
                            }
                        }
                        catch { }
                    });
                }
            }

            return task;
        }

        public Task UpdateTaskAsync(TaskItem task) => Task.CompletedTask;
        public async Task UpdateStatusAsync(int taskId, TaskItemStatus status, string? changedByUserId = null)
        {
            var existingTask = await _intra.GetTaskByIdAsync(taskId);
            var oldStatus = existingTask?.Status.ToString() ?? "Pending";
            await _intra.UpdateTaskStatusAsync(taskId, status);

            if (existingTask != null)
            {
                try
                {
                    var changerProfile = !string.IsNullOrEmpty(changedByUserId) ? await _intra.GetUserProfileAsync(changedByUserId) : null;
                    var changerName = changerProfile?.FullName ?? changedByUserId ?? "System";

                    var targetUsers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (!string.IsNullOrEmpty(existingTask.AssignedBy)) targetUsers.Add(existingTask.AssignedBy);
                    if (!string.IsNullOrEmpty(existingTask.AssignedTo)) targetUsers.Add(existingTask.AssignedTo);
                    if (!string.IsNullOrEmpty(existingTask.AdditionalAssigneeIds))
                    {
                        foreach (var a in existingTask.AdditionalAssigneeIds.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            var trimmed = a.Trim();
                            if (!string.IsNullOrEmpty(trimmed)) targetUsers.Add(trimmed);
                        }
                    }
                    if (!string.IsNullOrEmpty(changedByUserId)) targetUsers.Remove(changedByUserId);

                    foreach (var target in targetUsers)
                    {
                        _ = _notifications.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = $"Task Status: {existingTask.Title}",
                            Message = $"Status changed from {oldStatus} to {status} by {changerName}",
                            SenderName = changerName,
                            TargetUserId = target,
                            ActionUrl = "/office/tasks"
                        });
                    }
                }
                catch { }
            }
        }

        public Task DeleteTaskAsync(int id) => Task.CompletedTask;
        public async Task<TaskComment> AddCommentAsync(TaskComment comment)
        {
            var id = await _intra.AddTaskCommentAsync(comment);
            comment.Id = id;
            return comment;
        }

        public Task<List<TaskComment>> GetTaskCommentsAsync(int taskId) => Task.FromResult(new List<TaskComment>());
        public Task<TaskAttachment> AddAttachmentAsync(TaskAttachment attachment) => Task.FromResult(attachment);

        public async Task<int> GetUnreadTaskCountAsync(string userId)
        {
            var tasks = await _intra.GetTasksAsync(assignedTo: userId);
            return tasks.Count(t => t.Status == TaskItemStatus.Pending || t.Status == TaskItemStatus.InProgress);
        }

        public Task MarkTaskAsReadAsync(int taskId, string userId) => Task.CompletedTask;
        public Task MarkAllTasksAsReadAsync(string userId) => Task.CompletedTask;
    }

    // -------------------------------------------------------------
    // Meeting Service
    // -------------------------------------------------------------
    public class MeetingBroadcastOptions
    {
        public bool SendWhatsApp { get; set; } = true;
        public bool SendEmail { get; set; } = true;
        public bool SendInAppNotification { get; set; } = true;
        public bool EnableAutoMinutesDocx { get; set; } = true;
    }

    public interface IMeetingService
    {
        Task<Meeting> CreateMeetingAsync(Meeting meeting, List<string> participantUserIds, MeetingBroadcastOptions? broadcastOptions = null);
        Task<Meeting?> GetMeetingAsync(int meetingId);
        Task<List<Meeting>> GetUpcomingMeetingsAsync(string userId);
        Task<List<Meeting>> GetAllMeetingsForUserAsync(string userId);
        Task UpdateMeetingMinutesAsync(int meetingId, string minutes, string? voiceNotePath = null);
        Task<bool> GenerateAndDistributeMinutesAsync(int meetingId);
        Task<string?> GetOrCreateMinutesDocxPathAsync(int meetingId);
    }

    public class MeetingService : IMeetingService
    {
        private readonly IIntraOfficeService _intra;
        private readonly IAppNotificationService _notifications;

        public MeetingService(IIntraOfficeService intra, IAppNotificationService notifications)
        {
            _intra = intra;
            _notifications = notifications;
        }

        public async Task<Meeting> CreateMeetingAsync(Meeting meeting, List<string> participantUserIds, MeetingBroadcastOptions? broadcastOptions = null)
        {
            var id = await _intra.CreateMeetingAsync(meeting, participantUserIds);
            meeting.Id = id;

            var organizerProfile = !string.IsNullOrEmpty(meeting.OrganizerId) ? await _intra.GetUserProfileAsync(meeting.OrganizerId) : null;
            var organizerName = organizerProfile?.FullName ?? "Organizer";
            var timeString = meeting.ScheduledStartTime.ToLocalTime().ToString("MMM dd, yyyy hh:mm tt");

            foreach (var p in participantUserIds)
            {
                _ = _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Meeting,
                    Title = $"Meeting Invite: {meeting.Title}",
                    Message = $"Scheduled for {timeString} by {organizerName}",
                    SenderName = organizerName,
                    TargetUserId = p,
                    ActionUrl = "/office/meetings"
                });
            }
            return meeting;
        }

        public async Task<Meeting?> GetMeetingAsync(int meetingId) => await _intra.GetMeetingByIdAsync(meetingId);
        public async Task<List<Meeting>> GetUpcomingMeetingsAsync(string userId)
        {
            var all = await _intra.GetMeetingsForUserAsync(userId);
            return all.Where(m => m.ScheduledStartTime >= DateTime.UtcNow.AddHours(-2)).OrderBy(m => m.ScheduledStartTime).ToList();
        }

        public async Task<List<Meeting>> GetAllMeetingsForUserAsync(string userId) => await _intra.GetMeetingsForUserAsync(userId);
        public Task UpdateMeetingMinutesAsync(int meetingId, string minutes, string? voiceNotePath = null) => Task.CompletedTask;
        public Task<bool> GenerateAndDistributeMinutesAsync(int meetingId) => Task.FromResult(true);
        public Task<string?> GetOrCreateMinutesDocxPathAsync(int meetingId) => Task.FromResult<string?>(null);
    }

    // -------------------------------------------------------------
    // Minute Approval Service
    // -------------------------------------------------------------
    public interface IMinuteApprovalService
    {
        Task<MinuteApproval> CreateMinuteAsync(MinuteApproval minute, List<IBrowserFile>? attachments, IBrowserFile? signature);
        Task<List<MinuteApproval>> GetMyMinutesAsync(string userId);
        Task<List<MinuteApproval>> GetMinutesListAsync(MinuteFilter filter);
        Task<MinuteApproval?> GetMinuteByIdAsync(int id);
        Task ProcessWorkflowActionAsync(int minuteId, string userId, string actionTaken, string remarks, string? nextForwardToUserId, List<IBrowserFile>? attachments, IBrowserFile? signature);
        Task<List<MinuteWorkflowHistory>> GetMinuteHistoryAsync(int minuteId);
        Task<int> GetPendingApprovalsCountAsync(string userId);
        Task MarkAsReadAsync(int minuteId);
    }

    public class MinuteApprovalService : IMinuteApprovalService
    {
        private readonly IIntraOfficeService _intra;
        private readonly IAppNotificationService _notifications;

        public MinuteApprovalService(IIntraOfficeService intra, IAppNotificationService notifications)
        {
            _intra = intra;
            _notifications = notifications;
        }

        public async Task<MinuteApproval> CreateMinuteAsync(MinuteApproval minute, List<IBrowserFile>? attachments, IBrowserFile? signature)
        {
            var id = await _intra.CreateMinuteApprovalAsync(minute);
            minute.Id = id;
            if (!string.IsNullOrEmpty(minute.ForwardToUserId))
            {
                var creatorProfile = !string.IsNullOrEmpty(minute.CreatedByUserId) ? await _intra.GetUserProfileAsync(minute.CreatedByUserId) : null;
                var creatorName = creatorProfile?.FullName ?? minute.CreatedByUserId ?? "Colleague";
                _ = _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Minute,
                    Title = $"New Minute Pending Approval: #{minute.No}",
                    Message = $"{creatorName} forwarded minute: \"{minute.Subject}\"",
                    SenderName = creatorName,
                    TargetUserId = minute.ForwardToUserId,
                    ActionUrl = $"/office/minutes-workflow/{minute.Id}"
                });
            }
            return minute;
        }

        public async Task<List<MinuteApproval>> GetMyMinutesAsync(string userId) => await _intra.GetMinutesAsync(new MinuteFilter { InitiatorUserId = userId });
        public async Task<List<MinuteApproval>> GetMinutesListAsync(MinuteFilter filter) => await _intra.GetMinutesAsync(filter);
        public async Task<MinuteApproval?> GetMinuteByIdAsync(int id) => await _intra.GetMinuteByIdAsync(id);

        public async Task ProcessWorkflowActionAsync(int minuteId, string userId, string actionTaken, string remarks, string? nextForwardToUserId, List<IBrowserFile>? attachments, IBrowserFile? signature)
        {
            await _intra.UpdateMinuteStatusAsync(minuteId, actionTaken, userId, actionTaken, remarks);

            try
            {
                var minute = await _intra.GetMinuteByIdAsync(minuteId);
                var actorProfile = !string.IsNullOrEmpty(userId) ? await _intra.GetUserProfileAsync(userId) : null;
                var actorName = actorProfile?.FullName ?? userId ?? "Reviewer";

                // If referred or forwarded to next user
                if ((actionTaken == "Refer" || actionTaken == "Forward" || actionTaken == "Approve") && !string.IsNullOrEmpty(nextForwardToUserId))
                {
                    _ = _notifications.SendNotificationAsync(new AppNotification
                    {
                        Category = NotificationCategory.Minute,
                        Title = $"Minute {actionTaken}: #{minute?.No ?? minuteId.ToString()}",
                        Message = $"{actorName} forwarded minute: \"{minute?.Subject}\". Remarks: {remarks}",
                        SenderName = actorName,
                        TargetUserId = nextForwardToUserId,
                        ActionUrl = $"/office/minutes-workflow/{minuteId}"
                    });
                }

                // Notify original creator of status change
                if (minute != null && !string.IsNullOrEmpty(minute.CreatedByUserId) && !string.Equals(minute.CreatedByUserId, userId, StringComparison.OrdinalIgnoreCase))
                {
                    _ = _notifications.SendNotificationAsync(new AppNotification
                    {
                        Category = NotificationCategory.Minute,
                        Title = $"Minute #{minute.No} {actionTaken}",
                        Message = $"{actorName} updated status to {actionTaken}. Remarks: {remarks}",
                        SenderName = actorName,
                        TargetUserId = minute.CreatedByUserId,
                        ActionUrl = $"/office/minutes-workflow/{minuteId}"
                    });
                }
            }
            catch { }
        }

        public Task<List<MinuteWorkflowHistory>> GetMinuteHistoryAsync(int minuteId) => Task.FromResult(new List<MinuteWorkflowHistory>());
        public async Task<int> GetPendingApprovalsCountAsync(string userId)
        {
            var list = await _intra.GetMinutesAsync(new MinuteFilter { ForwardToUserId = userId, Status = "Pending" });
            return list.Count;
        }
        public Task MarkAsReadAsync(int minuteId) => Task.CompletedTask;
    }

    // -------------------------------------------------------------
    // WhatsApp Service
    // -------------------------------------------------------------
    public interface IWhatsAppService
    {
        Task<bool> SendTaskNotificationAsync(string phoneNumber, string taskTitle, string assigneeName, string priority, List<string>? absoluteAttachmentPaths = null, string? description = null);
        Task<bool> SendDirectMessageAsync(string phoneNumber, string message, List<string>? absoluteAttachmentPaths = null);
        Task<bool> SendMinuteApprovalRequestAsync(string phoneNumber, string minuteTitle, string approverName, string submitterName, string link);
    }

    public class WhatsAppService : IWhatsAppService
    {
        private readonly IWhatsAppNotificationService _client;
        public WhatsAppService(IWhatsAppNotificationService client) => _client = client;

        public Task<bool> SendTaskNotificationAsync(string phoneNumber, string taskTitle, string assigneeName, string priority, List<string>? absoluteAttachmentPaths = null, string? description = null)
        {
            return _client.SendTaskNotificationAsync(phoneNumber, taskTitle, assigneeName, priority, description, absoluteAttachmentPaths);
        }

        public Task<bool> SendDirectMessageAsync(string phoneNumber, string message, List<string>? absoluteAttachmentPaths = null)
            => _client.SendDirectMessageAsync(phoneNumber, message, absoluteAttachmentPaths);

        public Task<bool> SendMinuteApprovalRequestAsync(string phoneNumber, string minuteTitle, string approverName, string submitterName, string link)
        {
            var msg = $"*Executive Minute Approval Request*\n*Title:* {minuteTitle}\n*For:* {approverName}\n*From:* {submitterName}\n*Review:* {link}";
            return _client.SendDirectMessageAsync(phoneNumber, msg);
        }
    }

    // -------------------------------------------------------------
    // Email Service
    // -------------------------------------------------------------
    public interface IEmailService
    {
        Task<bool> SendEmailAsync(string toEmail, string subject, string body, List<string>? attachmentPaths = null);
        Task<(bool Success, string Message)> TestEmailConfigurationAsync(string recipientEmail, EmailConfiguration config, string plainPassword);
    }

    public class EmailService : IEmailService
    {
        private readonly IIntraOfficeService _intra;
        private readonly IEmailEncryptionService _encryption;
        private readonly ILogger<EmailService> _logger;

        public EmailService(IIntraOfficeService intra, IEmailEncryptionService encryption, ILogger<EmailService> logger)
        {
            _intra = intra;
            _encryption = encryption;
            _logger = logger;
        }

        public async Task<bool> SendEmailAsync(string toEmail, string subject, string body, List<string>? attachmentPaths = null)
        {
            try
            {
                var config = await _intra.GetEmailConfigurationAsync();
                if (config == null || string.IsNullOrWhiteSpace(config.SenderEmail))
                {
                    _logger.LogWarning("Email configuration is not set up.");
                    return false;
                }

                var plainPassword = _encryption.Decrypt(config.EncryptedPassword);
                using var client = new System.Net.Mail.SmtpClient(config.SmtpServer, config.SmtpPort)
                {
                    EnableSsl = config.EnableSsl,
                    Credentials = new System.Net.NetworkCredential(config.SenderEmail, plainPassword),
                    Timeout = 15000
                };

                using var msg = new System.Net.Mail.MailMessage
                {
                    From = new System.Net.Mail.MailAddress(config.SenderEmail, string.IsNullOrWhiteSpace(config.SenderName) ? "IntraCom-CRM" : config.SenderName),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = body.Contains("<html") || body.Contains("<div") || body.Contains("<p")
                };
                msg.To.Add(toEmail);

                if (attachmentPaths != null)
                {
                    foreach (var path in attachmentPaths.Where(p => System.IO.File.Exists(p)))
                    {
                        msg.Attachments.Add(new System.Net.Mail.Attachment(path));
                    }
                }

                await client.SendMailAsync(msg);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error sending email to {ToEmail}", toEmail);
                return false;
            }
        }

        public async Task<(bool Success, string Message)> TestEmailConfigurationAsync(string recipientEmail, EmailConfiguration config, string plainPassword)
        {
            try
            {
                using var client = new System.Net.Mail.SmtpClient(config.SmtpServer, config.SmtpPort)
                {
                    EnableSsl = config.EnableSsl,
                    Credentials = new System.Net.NetworkCredential(config.SenderEmail, plainPassword),
                    Timeout = 12000
                };

                using var msg = new System.Net.Mail.MailMessage
                {
                    From = new System.Net.Mail.MailAddress(config.SenderEmail, string.IsNullOrWhiteSpace(config.SenderName) ? "IntraCom-CRM" : config.SenderName),
                    Subject = "IntraCom-CRM SMTP Connection Test",
                    Body = "This is an automated test message from IntraCom-CRM verifying that your SMTP server settings and credentials are configured correctly."
                };
                msg.To.Add(recipientEmail);

                await client.SendMailAsync(msg);
                return (true, $"Test email sent successfully to {recipientEmail}!");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "SMTP connection test failed");
                return (false, "SMTP connection test failed: " + ex.Message);
            }
        }
    }

    // -------------------------------------------------------------
    // Minute Type Service
    // -------------------------------------------------------------
    public interface IMinuteTypeService
    {
        Task<List<MinuteType>> GetAllAsync();
        Task<List<MinuteType>> GetActiveAsync();
        Task<MinuteType?> GetByIdAsync(int id);
        Task<MinuteType> CreateAsync(string name);
        Task UpdateAsync(int id, string name, bool isActive);
        Task DeleteAsync(int id);
    }

    public class MinuteTypeService : IMinuteTypeService
    {
        private readonly IIntraOfficeService _intra;
        public MinuteTypeService(IIntraOfficeService intra) => _intra = intra;

        public async Task<List<MinuteType>> GetAllAsync() => await _intra.GetAllMinuteTypesAsync();
        public async Task<List<MinuteType>> GetActiveAsync()
        {
            var all = await _intra.GetMinuteTypesAsync();
            return all.Where(t => t.IsActive).OrderBy(t => t.Name).ToList();
        }
        public async Task<MinuteType?> GetByIdAsync(int id)
        {
            var all = await _intra.GetAllMinuteTypesAsync();
            return all.FirstOrDefault(t => t.Id == id);
        }
        public async Task<MinuteType> CreateAsync(string name)
        {
            var id = await _intra.CreateMinuteTypeAsync(name);
            return new MinuteType { Id = id, Name = name, IsActive = true, CreatedAt = DateTime.UtcNow };
        }
        public async Task UpdateAsync(int id, string name, bool isActive) => await _intra.UpdateMinuteTypeAsync(id, name, isActive);
        public async Task DeleteAsync(int id) => await _intra.DeleteMinuteTypeAsync(id);
    }

    // -------------------------------------------------------------
    // CRM Service (Reports & Templates)
    // -------------------------------------------------------------
    public interface ICrmService
    {
        Task<DashboardMetricsDto> GetDashboardMetricsAsync();
        Task<ARAgingSummaryDto> GetARAgingSummaryAsync();
        Task<List<EmailTemplate>> GetEmailTemplatesAsync(string? category = null);
        Task<EmailTemplate?> GetEmailTemplateByCodeAsync(string code);
        Task<EmailTemplate> SaveEmailTemplateAsync(EmailTemplate template, string userName);
        (string Subject, string Body) RenderEmailTemplate(EmailTemplate template, Dictionary<string, string> placeholders);
        Task SaveActivityAsync(LeadActivityModel activity, string userName);
        Task<List<LeadModel>> GetCustomersAsync();
    }

    public class CrmService : ICrmService
    {
        private readonly IIntraOfficeService _intra;
        public CrmService(IIntraOfficeService intra) => _intra = intra;

        public Task<DashboardMetricsDto> GetDashboardMetricsAsync() => _intra.GetDashboardMetricsAsync();
        public Task<ARAgingSummaryDto> GetARAgingSummaryAsync() => _intra.GetARAgingSummaryAsync();
        public Task<List<EmailTemplate>> GetEmailTemplatesAsync(string? category = null) => _intra.GetEmailTemplatesAsync(category);
        public Task<EmailTemplate?> GetEmailTemplateByCodeAsync(string code) => _intra.GetEmailTemplateByCodeAsync(code);

        public async Task<EmailTemplate> SaveEmailTemplateAsync(EmailTemplate template, string userName)
        {
            var id = await _intra.SaveEmailTemplateAsync(template);
            template.Id = id;
            return template;
        }

        public (string Subject, string Body) RenderEmailTemplate(EmailTemplate template, Dictionary<string, string> placeholders)
        {
            var subject = template.SubjectTemplate;
            var body = template.BodyTemplate;

            foreach (var kvp in placeholders)
            {
                subject = subject.Replace(kvp.Key, kvp.Value);
                body = body.Replace(kvp.Key, kvp.Value);
            }

            return (subject, body);
        }

        public Task SaveActivityAsync(LeadActivityModel activity, string userName) => _intra.AddLeadActivityAsync(activity);
        public Task<List<LeadModel>> GetCustomersAsync() => _intra.GetLeadsAsync();
    }
}

