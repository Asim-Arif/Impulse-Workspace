using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;

namespace DataAccessLibrary.Interface.IntraOffice
{
    public interface IIntraOfficeDataAccess
    {
        // 1. Users & Directory
        Task<List<IntraUserProfile>> GetActiveUsersAsync(string? search = null);
        Task<IntraUserProfile?> GetUserProfileAsync(string userId);
        Task<List<(string DeptId, string DeptName)>> GetDepartmentsAsync();

        // 2. Channels
        Task<List<Channel>> GetChannelsForUserAsync(string userId);
        Task<Channel?> GetChannelByIdAsync(int channelId);
        Task<int> CreateChannelAsync(Channel channel, List<string>? memberUserIds = null);
        Task<bool> AddMemberToChannelAsync(int channelId, string userId, bool isAdmin = false);
        Task<bool> RemoveMemberFromChannelAsync(int channelId, string userId);

        // 3. Messages & Direct Chat
        Task<List<Message>> GetChannelMessagesAsync(int channelId, int take = 50, long? beforeMessageId = null);
        Task<List<Message>> GetDirectMessagesAsync(string userId1, string userId2, int take = 50, long? beforeMessageId = null);
        Task<List<DirectChatSummary>> GetDirectChatSummariesAsync(string currentUserId);
        Task<long> SendMessageAsync(Message message);
        Task<bool> MarkMessagesAsReadAsync(string currentUserId, string? senderId = null, int? channelId = null);
        Task<bool> DeleteMessageAsync(long messageId, string userId);

        // 4. Announcements
        Task<List<Announcement>> GetAnnouncementsAsync(string currentUserId, string? departmentId = null);
        Task<Announcement?> GetAnnouncementByIdAsync(int id, string currentUserId);
        Task<int> CreateAnnouncementAsync(Announcement announcement);
        Task<bool> AcknowledgeAnnouncementAsync(int announcementId, string userId);

        // 5. Tasks
        Task<List<TaskItem>> GetTasksAsync(string? assignedTo = null, string? assignedBy = null, string? departmentId = null, TaskItemStatus? status = null);
        Task<TaskItem?> GetTaskByIdAsync(int taskId);
        Task<int> CreateTaskAsync(TaskItem task);
        Task<bool> UpdateTaskStatusAsync(int taskId, TaskItemStatus status, DateTime? completedAt = null);
        Task<int> AddTaskCommentAsync(TaskComment comment);
        Task<List<TaskDueMonitoringDto>> GetTasksPending70PercentDueWarningAsync();
        Task<List<TaskDueMonitoringDto>> GetTasksPendingOverdueAlertAsync();
        Task<bool> MarkDueWarningSentAsync(int taskId);
        Task<bool> MarkOverdueWarningSentAsync(int taskId);
        Task<List<string>> GetHubSupervisorsForLotTaskAsync(string lotNo, string hubName);

        // 6. Meetings
        Task<List<Meeting>> GetMeetingsForUserAsync(string userId);
        Task<Meeting?> GetMeetingByIdAsync(int meetingId);
        Task<int> CreateMeetingAsync(Meeting meeting, List<string>? participantUserIds = null);
        Task<bool> UpdateMeetingStatusAsync(int meetingId, MeetingStatus status);

        // 7. Minutes Approval
        Task<List<MinuteApproval>> GetMinutesAsync(MinuteFilter filter);
        Task<MinuteApproval?> GetMinuteByIdAsync(int id);
        Task<int> CreateMinuteApprovalAsync(MinuteApproval minute);
        Task<bool> UpdateMinuteStatusAsync(int minuteId, string status, string userId, string actionTaken, string? remarks = null, string? signaturePath = null);
        Task<List<MinuteType>> GetMinuteTypesAsync();
        Task<List<MinuteType>> GetAllMinuteTypesAsync();
        Task<int> CreateMinuteTypeAsync(string name);
        Task<bool> UpdateMinuteTypeAsync(int id, string name, bool isActive);
        Task<bool> DeleteMinuteTypeAsync(int id);

        // 8. Presence & Sticky Notes
        Task<bool> UpdateUserPresenceAsync(string userId, UserStatus status, string? connectionId = null);
        Task<List<UserPresence>> GetUserPresencesAsync();
        Task<List<StickyNote>> GetStickyNotesAsync(string userId);
        Task<int> SaveStickyNoteAsync(StickyNote note);
        Task<bool> DeleteStickyNoteAsync(int id, string userId);

        // 9. Leads & Pipeline
        Task<List<LeadModel>> GetLeadsAsync(string? status = null, string? priority = null, string? search = null);
        Task<LeadModel?> GetLeadByIdAsync(int id);
        Task<int> SaveLeadAsync(LeadModel lead);
        Task<bool> DeleteLeadAsync(int id);
        Task<List<LeadActivityModel>> GetLeadActivitiesAsync(int leadId);
        Task<int> AddLeadActivityAsync(LeadActivityModel activity);
        Task<bool> ConvertLeadToForeignCustomerAsync(ConvertLeadToCustomerModel model);
        Task<bool> CheckCustCodeExistsAsync(string custCode);

        // 10. Customer 360 Hub
        Task<Customer360Dto> GetCustomer360Async(string? customerCode = null);
        Task<List<Customer360LookupDto>> GetCustomer360LookupListAsync();
        Task<bool> ConvertCustomer360ToForeignCustomerAsync(ConvertCustomerToForeignCustomerModel model);
        Task<bool> AddCustomer360ActivityAsync(CustomerActivityDto activity);
        Task<bool> AddCustomer360ContactAsync(CustomerContactDto contact);

        // 11. Email & SMTP Configuration
        Task<EmailConfiguration?> GetEmailConfigurationAsync();
        Task<bool> SaveEmailConfigurationAsync(EmailConfiguration config);

        // 12. Email Templates
        Task<List<EmailTemplate>> GetEmailTemplatesAsync(string? category = null);
        Task<EmailTemplate?> GetEmailTemplateByCodeAsync(string code);
        Task<EmailTemplate?> GetEmailTemplateByIdAsync(int id);
        Task<int> SaveEmailTemplateAsync(EmailTemplate template);
        Task<bool> DeleteEmailTemplateAsync(int id);

        // 13. Executive Reports & Business Intelligence
        Task<DashboardMetricsDto> GetDashboardMetricsAsync();
        Task<ARAgingSummaryDto> GetARAgingSummaryAsync();
    }
}

