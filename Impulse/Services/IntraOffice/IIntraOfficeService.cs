using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;

namespace Impulse.Services.IntraOffice
{
    public interface IIntraOfficeService
    {
        // Directory & Users
        Task<List<IntraUserProfile>> GetActiveUsersAsync(string? search = null);
        Task<IntraUserProfile?> GetUserProfileAsync(string userId);
        Task<List<(string DeptId, string DeptName)>> GetDepartmentsAsync();

        // Channels
        Task<List<Channel>> GetChannelsForUserAsync(string userId);
        Task<Channel?> GetChannelByIdAsync(int channelId);
        Task<int> CreateChannelAsync(Channel channel, List<string>? memberUserIds = null);
        Task<bool> AddMemberToChannelAsync(int channelId, string userId, bool isAdmin = false);
        Task<bool> RemoveMemberFromChannelAsync(int channelId, string userId);

        // Messages & Direct Chat
        Task<List<Message>> GetChannelMessagesAsync(int channelId, int take = 50, long? beforeMessageId = null);
        Task<List<Message>> GetDirectMessagesAsync(string userId1, string userId2, int take = 50, long? beforeMessageId = null);
        Task<List<DirectChatSummary>> GetDirectChatSummariesAsync(string currentUserId);
        Task<long> SendMessageAsync(Message message);
        Task<bool> MarkMessagesAsReadAsync(string currentUserId, string? senderId = null, int? channelId = null);
        Task<bool> DeleteMessageAsync(long messageId, string userId);

        // Announcements
        Task<List<Announcement>> GetAnnouncementsAsync(string currentUserId, string? departmentId = null);
        Task<Announcement?> GetAnnouncementByIdAsync(int id, string currentUserId);
        Task<int> CreateAnnouncementAsync(Announcement announcement);
        Task<bool> AcknowledgeAnnouncementAsync(int announcementId, string userId);

        // Tasks
        Task<List<TaskItem>> GetTasksAsync(string? assignedTo = null, string? assignedBy = null, string? departmentId = null, TaskItemStatus? status = null);
        Task<TaskItem?> GetTaskByIdAsync(int taskId);
        Task<int> CreateTaskAsync(TaskItem task);
        Task<bool> UpdateTaskStatusAsync(int taskId, TaskItemStatus status);
        Task<int> AddTaskCommentAsync(TaskComment comment);

        // Meetings
        Task<List<Meeting>> GetMeetingsForUserAsync(string userId);
        Task<Meeting?> GetMeetingByIdAsync(int meetingId);
        Task<int> CreateMeetingAsync(Meeting meeting, List<string>? participantUserIds = null);
        Task<bool> UpdateMeetingStatusAsync(int meetingId, MeetingStatus status);

        // Minutes Approval
        Task<List<MinuteApproval>> GetMinutesAsync(MinuteFilter filter);
        Task<MinuteApproval?> GetMinuteByIdAsync(int id);
        Task<int> CreateMinuteApprovalAsync(MinuteApproval minute);
        Task<bool> UpdateMinuteStatusAsync(int minuteId, string status, string userId, string actionTaken, string? remarks = null, string? signaturePath = null);
        Task<List<MinuteType>> GetMinuteTypesAsync();

        // Presence & Sticky Notes
        Task<bool> UpdateUserPresenceAsync(string userId, UserStatus status, string? connectionId = null);
        Task<List<UserPresence>> GetUserPresencesAsync();
        Task<List<StickyNote>> GetStickyNotesAsync(string userId);
        Task<int> SaveStickyNoteAsync(StickyNote note);
        Task<bool> DeleteStickyNoteAsync(int id, string userId);

        // Leads & Pre-Sales Pipeline
        Task<List<LeadModel>> GetLeadsAsync(string? status = null, string? priority = null, string? search = null);
        Task<LeadModel?> GetLeadByIdAsync(int id);
        Task<int> SaveLeadAsync(LeadModel lead);
        Task<bool> DeleteLeadAsync(int id);
        Task<List<LeadActivityModel>> GetLeadActivitiesAsync(int leadId);
        Task<int> AddLeadActivityAsync(LeadActivityModel activity);
        Task<bool> ConvertLeadToForeignCustomerAsync(ConvertLeadToCustomerModel model);
        Task<bool> CheckCustCodeExistsAsync(string custCode);
    }
}
