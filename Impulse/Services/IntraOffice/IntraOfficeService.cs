using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.IntraOffice;
using DataAccessLibrary.Models.IntraOffice;

namespace Impulse.Services.IntraOffice
{
    public class IntraOfficeService : IIntraOfficeService
    {
        private readonly IIntraOfficeDataAccess _dataAccess;

        public IntraOfficeService(IIntraOfficeDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<IntraUserProfile>> GetActiveUsersAsync(string? search = null) => _dataAccess.GetActiveUsersAsync(search);
        public Task<IntraUserProfile?> GetUserProfileAsync(string userId) => _dataAccess.GetUserProfileAsync(userId);
        public Task<List<(string DeptId, string DeptName)>> GetDepartmentsAsync() => _dataAccess.GetDepartmentsAsync();

        public Task<List<Channel>> GetChannelsForUserAsync(string userId) => _dataAccess.GetChannelsForUserAsync(userId);
        public Task<Channel?> GetChannelByIdAsync(int channelId) => _dataAccess.GetChannelByIdAsync(channelId);
        public Task<int> CreateChannelAsync(Channel channel, List<string>? memberUserIds = null) => _dataAccess.CreateChannelAsync(channel, memberUserIds);
        public Task<bool> AddMemberToChannelAsync(int channelId, string userId, bool isAdmin = false) => _dataAccess.AddMemberToChannelAsync(channelId, userId, isAdmin);
        public Task<bool> RemoveMemberFromChannelAsync(int channelId, string userId) => _dataAccess.RemoveMemberFromChannelAsync(channelId, userId);

        public Task<List<Message>> GetChannelMessagesAsync(int channelId, int take = 50, long? beforeMessageId = null) => _dataAccess.GetChannelMessagesAsync(channelId, take, beforeMessageId);
        public Task<List<Message>> GetDirectMessagesAsync(string userId1, string userId2, int take = 50, long? beforeMessageId = null) => _dataAccess.GetDirectMessagesAsync(userId1, userId2, take, beforeMessageId);
        public Task<List<DirectChatSummary>> GetDirectChatSummariesAsync(string currentUserId) => _dataAccess.GetDirectChatSummariesAsync(currentUserId);
        public Task<long> SendMessageAsync(Message message) => _dataAccess.SendMessageAsync(message);
        public Task<bool> MarkMessagesAsReadAsync(string currentUserId, string? senderId = null, int? channelId = null) => _dataAccess.MarkMessagesAsReadAsync(currentUserId, senderId, channelId);
        public Task<bool> DeleteMessageAsync(long messageId, string userId) => _dataAccess.DeleteMessageAsync(messageId, userId);

        public Task<List<Announcement>> GetAnnouncementsAsync(string currentUserId, string? departmentId = null) => _dataAccess.GetAnnouncementsAsync(currentUserId, departmentId);
        public Task<Announcement?> GetAnnouncementByIdAsync(int id, string currentUserId) => _dataAccess.GetAnnouncementByIdAsync(id, currentUserId);
        public Task<int> CreateAnnouncementAsync(Announcement announcement) => _dataAccess.CreateAnnouncementAsync(announcement);
        public Task<bool> AcknowledgeAnnouncementAsync(int announcementId, string userId) => _dataAccess.AcknowledgeAnnouncementAsync(announcementId, userId);

        public Task<List<TaskItem>> GetTasksAsync(string? assignedTo = null, string? assignedBy = null, string? departmentId = null, TaskItemStatus? status = null) => _dataAccess.GetTasksAsync(assignedTo, assignedBy, departmentId, status);
        public Task<TaskItem?> GetTaskByIdAsync(int taskId) => _dataAccess.GetTaskByIdAsync(taskId);
        public Task<int> CreateTaskAsync(TaskItem task) => _dataAccess.CreateTaskAsync(task);
        public Task<bool> UpdateTaskStatusAsync(int taskId, TaskItemStatus status) => _dataAccess.UpdateTaskStatusAsync(taskId, status);
        public Task<int> AddTaskCommentAsync(TaskComment comment) => _dataAccess.AddTaskCommentAsync(comment);

        public Task<List<Meeting>> GetMeetingsForUserAsync(string userId) => _dataAccess.GetMeetingsForUserAsync(userId);
        public Task<Meeting?> GetMeetingByIdAsync(int meetingId) => _dataAccess.GetMeetingByIdAsync(meetingId);
        public Task<int> CreateMeetingAsync(Meeting meeting, List<string>? participantUserIds = null) => _dataAccess.CreateMeetingAsync(meeting, participantUserIds);
        public Task<bool> UpdateMeetingStatusAsync(int meetingId, MeetingStatus status) => _dataAccess.UpdateMeetingStatusAsync(meetingId, status);

        public Task<List<MinuteApproval>> GetMinutesAsync(MinuteFilter filter) => _dataAccess.GetMinutesAsync(filter);
        public Task<MinuteApproval?> GetMinuteByIdAsync(int id) => _dataAccess.GetMinuteByIdAsync(id);
        public Task<int> CreateMinuteApprovalAsync(MinuteApproval minute) => _dataAccess.CreateMinuteApprovalAsync(minute);
        public Task<bool> UpdateMinuteStatusAsync(int minuteId, string status, string userId, string actionTaken, string? remarks = null, string? signaturePath = null) => _dataAccess.UpdateMinuteStatusAsync(minuteId, status, userId, actionTaken, remarks, signaturePath);
        public Task<List<MinuteType>> GetMinuteTypesAsync() => _dataAccess.GetMinuteTypesAsync();

        public Task<bool> UpdateUserPresenceAsync(string userId, UserStatus status, string? connectionId = null) => _dataAccess.UpdateUserPresenceAsync(userId, status, connectionId);
        public Task<List<UserPresence>> GetUserPresencesAsync() => _dataAccess.GetUserPresencesAsync();
        public Task<List<StickyNote>> GetStickyNotesAsync(string userId) => _dataAccess.GetStickyNotesAsync(userId);
        public Task<int> SaveStickyNoteAsync(StickyNote note) => _dataAccess.SaveStickyNoteAsync(note);
        public Task<bool> DeleteStickyNoteAsync(int id, string userId) => _dataAccess.DeleteStickyNoteAsync(id, userId);

        // Leads & CRM Integration
        public Task<List<LeadModel>> GetLeadsAsync(string? status = null, string? source = null, string? search = null) => _dataAccess.GetLeadsAsync(status, source, search);
        public Task<LeadModel?> GetLeadByIdAsync(int id) => _dataAccess.GetLeadByIdAsync(id);
        public Task<int> SaveLeadAsync(LeadModel lead) => _dataAccess.SaveLeadAsync(lead);
        public Task<bool> DeleteLeadAsync(int id) => _dataAccess.DeleteLeadAsync(id);
        public Task<List<LeadActivityModel>> GetLeadActivitiesAsync(int leadId) => _dataAccess.GetLeadActivitiesAsync(leadId);
        public Task<int> AddLeadActivityAsync(LeadActivityModel activity) => _dataAccess.AddLeadActivityAsync(activity);
        public Task<bool> CheckCustCodeExistsAsync(string custCode) => _dataAccess.CheckCustCodeExistsAsync(custCode);
        public Task<bool> ConvertLeadToForeignCustomerAsync(ConvertLeadToCustomerModel model) => _dataAccess.ConvertLeadToForeignCustomerAsync(model);

        // Customer 360 Hub
        public Task<Customer360Dto> GetCustomer360Async(string? customerCode = null) => _dataAccess.GetCustomer360Async(customerCode);
        public Task<List<Customer360LookupDto>> GetCustomer360LookupListAsync() => _dataAccess.GetCustomer360LookupListAsync();
        public Task<bool> ConvertCustomer360ToForeignCustomerAsync(ConvertCustomerToForeignCustomerModel model) => _dataAccess.ConvertCustomer360ToForeignCustomerAsync(model);
        public Task<bool> AddCustomer360ActivityAsync(CustomerActivityDto activity) => _dataAccess.AddCustomer360ActivityAsync(activity);
        public Task<bool> AddCustomer360ContactAsync(CustomerContactDto contact) => _dataAccess.AddCustomer360ContactAsync(contact);

        // Minute Types Admin
        public Task<List<MinuteType>> GetAllMinuteTypesAsync() => _dataAccess.GetAllMinuteTypesAsync();
        public Task<int> CreateMinuteTypeAsync(string name) => _dataAccess.CreateMinuteTypeAsync(name);
        public Task<bool> UpdateMinuteTypeAsync(int id, string name, bool isActive) => _dataAccess.UpdateMinuteTypeAsync(id, name, isActive);
        public Task<bool> DeleteMinuteTypeAsync(int id) => _dataAccess.DeleteMinuteTypeAsync(id);

        // Email Configuration
        public Task<EmailConfiguration?> GetEmailConfigurationAsync() => _dataAccess.GetEmailConfigurationAsync();
        public Task<bool> SaveEmailConfigurationAsync(EmailConfiguration config) => _dataAccess.SaveEmailConfigurationAsync(config);

        // Email Templates
        public Task<List<EmailTemplate>> GetEmailTemplatesAsync(string? category = null) => _dataAccess.GetEmailTemplatesAsync(category);
        public Task<EmailTemplate?> GetEmailTemplateByCodeAsync(string code) => _dataAccess.GetEmailTemplateByCodeAsync(code);
        public Task<EmailTemplate?> GetEmailTemplateByIdAsync(int id) => _dataAccess.GetEmailTemplateByIdAsync(id);
        public Task<int> SaveEmailTemplateAsync(EmailTemplate template) => _dataAccess.SaveEmailTemplateAsync(template);
        public Task<bool> DeleteEmailTemplateAsync(int id) => _dataAccess.DeleteEmailTemplateAsync(id);

        // Executive Reports & Analytics
        public Task<DashboardMetricsDto> GetDashboardMetricsAsync() => _dataAccess.GetDashboardMetricsAsync();
        public Task<ARAgingSummaryDto> GetARAgingSummaryAsync() => _dataAccess.GetARAgingSummaryAsync();
    }
}

