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

            if (!string.IsNullOrEmpty(message.ReceiverId))
            {
                await _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Message,
                    Title = $"Message from {message.SenderName ?? message.SenderId}",
                    Message = message.Content,
                    SenderName = message.SenderName ?? message.SenderId,
                    TargetUserId = message.ReceiverId,
                    ActionUrl = $"/office/messages/{message.SenderId}"
                });
            }
            return message;
        }

        public Task<Message?> GetMessageByIdAsync(long id) => Task.FromResult<Message?>(null);
        public Task MarkAsReadAsync(long messageId) => Task.CompletedTask;
        public async Task MarkAllAsReadAsync(string userId, int? channelId = null, string? senderId = null)
            => await _intra.MarkMessagesAsReadAsync(userId, senderId, channelId);

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
            await _notifications.SendNotificationAsync(new AppNotification
            {
                Category = NotificationCategory.Announcement,
                Title = announcement.Title,
                Message = announcement.Content,
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
            => await _intra.AcknowledgeAnnouncementAsync(announcementId, userId);

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
            var id = await _intra.CreateTaskAsync(task);
            task.Id = id;

            if (!string.IsNullOrEmpty(task.AssignedTo))
            {
                await _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Task,
                    Title = $"Task Assigned: {task.Title}",
                    Message = task.Description ?? "You have been assigned a new task.",
                    TargetUserId = task.AssignedTo,
                    ActionUrl = "/office/tasks"
                });

                if (sendWhatsApp)
                {
                    var userProfile = await _intra.GetUserProfileAsync(task.AssignedTo);
                    if (!string.IsNullOrEmpty(userProfile?.CellNo))
                    {
                        task.WhatsAppMessageSent = await _whatsApp.SendTaskNotificationAsync(userProfile.CellNo, task.Title, userProfile.FullName ?? task.AssignedTo, task.Priority.ToString(), null, task.Description);
                    }
                }
            }

            return task;
        }

        public Task UpdateTaskAsync(TaskItem task) => Task.CompletedTask;
        public async Task UpdateStatusAsync(int taskId, TaskItemStatus status, string? changedByUserId = null)
            => await _intra.UpdateTaskStatusAsync(taskId, status);

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
            foreach (var p in participantUserIds)
            {
                await _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Meeting,
                    Title = $"Meeting Invitation: {meeting.Title}",
                    Message = $"Starts at {meeting.ScheduledStartTime:g}",
                    TargetUserId = p,
                    ActionUrl = $"/office/meetings/live/{id}"
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
                await _notifications.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Minute,
                    Title = $"Minute Approval: {minute.Title}",
                    Message = minute.Subject,
                    TargetUserId = minute.ForwardToUserId,
                    ActionUrl = $"/office/minutes-list"
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
    }

    public class EmailService : IEmailService
    {
        public Task<bool> SendEmailAsync(string toEmail, string subject, string body, List<string>? attachmentPaths = null)
        {
            // Background email service placeholder
            return Task.FromResult(true);
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

        public async Task<List<MinuteType>> GetAllAsync() => await _intra.GetMinuteTypesAsync();
        public async Task<List<MinuteType>> GetActiveAsync()
        {
            var all = await _intra.GetMinuteTypesAsync();
            return all.Where(t => t.IsActive).OrderBy(t => t.Name).ToList();
        }
        public async Task<MinuteType?> GetByIdAsync(int id)
        {
            var all = await _intra.GetMinuteTypesAsync();
            return all.FirstOrDefault(t => t.Id == id);
        }
        public Task<MinuteType> CreateAsync(string name) => Task.FromResult(new MinuteType { Name = name, IsActive = true });
        public Task UpdateAsync(int id, string name, bool isActive) => Task.CompletedTask;
        public Task DeleteAsync(int id) => Task.CompletedTask;
    }
}
