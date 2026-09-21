using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.IntraOffice;
using DataAccessLibrary.Models.IntraOffice;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.IntraOffice
{
    public class IntraOfficeDataAccess : IIntraOfficeDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<IntraOfficeDataAccess> _logger;

        public IntraOfficeDataAccess(IConfiguration configuration, ILogger<IntraOfficeDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        #region 1. Users & Directory

        public async Task<List<IntraUserProfile>> GetActiveUsersAsync(string? search = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        u.UserID,
                        u.UserName,
                        COALESCE(u.FullUserName, e.name, u.UserName) AS FullUserName,
                        u.EmpID,
                        e.name AS EmployeeName,
                        e.Designation,
                        e.deptid AS DepartmentId,
                        d.name AS DepartmentName,
                        e.Phone1 AS CellNo,
                        COALESCE(p.Status, 0) AS Status,
                        COALESCE(p.LastSeen, GETUTCDATE()) AS LastSeen,
                        1 AS IsActive
                    FROM Users u
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    LEFT JOIN UserPresences p ON u.UserName = p.UserId
                    WHERE (@Search IS NULL OR u.UserName LIKE @SearchParam OR e.name LIKE @SearchParam OR e.Designation LIKE @SearchParam OR d.name LIKE @SearchParam)
                    ORDER BY u.UserName";

                var searchParam = string.IsNullOrWhiteSpace(search) ? null : $"%{search.Trim()}%";
                var result = await db.QueryAsync<IntraUserProfile>(sql, new { Search = search, SearchParam = searchParam });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting active users for IntraOffice");
                return new List<IntraUserProfile>();
            }
        }

        public async Task<IntraUserProfile?> GetUserProfileAsync(string userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        u.UserID,
                        u.UserName,
                        COALESCE(u.FullUserName, e.name, u.UserName) AS FullUserName,
                        u.EmpID,
                        e.name AS EmployeeName,
                        e.Designation,
                        e.deptid AS DepartmentId,
                        d.name AS DepartmentName,
                        e.Phone1 AS CellNo,
                        COALESCE(p.Status, 0) AS Status,
                        COALESCE(p.LastSeen, GETUTCDATE()) AS LastSeen,
                        1 AS IsActive
                    FROM Users u
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    LEFT JOIN UserPresences p ON u.UserName = p.UserId
                    WHERE u.UserName = @UserId OR u.EmpID = @UserId";

                return await db.QueryFirstOrDefaultAsync<IntraUserProfile>(sql, new { UserId = userId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting user profile for {UserId}", userId);
                return null;
            }
        }

        public async Task<List<(string DeptId, string DeptName)>> GetDepartmentsAsync()
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT deptid AS DeptId, name AS DeptName FROM Departments WHERE Active = 1 ORDER BY name";
                var rows = await db.QueryAsync(sql);
                return rows.Select(r => ((string)r.DeptId, (string)r.DeptName)).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting departments");
                return new List<(string, string)>();
            }
        }

        #endregion

        #region 2. Channels

        public async Task<List<Channel>> GetChannelsForUserAsync(string userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        c.Id,
                        c.Name,
                        c.Description,
                        c.DepartmentId,
                        d.name AS DepartmentName,
                        c.IsPrivate,
                        c.IsActive,
                        c.CreatedBy,
                        u.FullUserName AS CreatorName,
                        c.CreatedAt,
                        (SELECT COUNT(1) FROM ChannelMembers cm WHERE cm.ChannelId = c.Id) AS MemberCount,
                        (SELECT COUNT(1) FROM Messages m WHERE m.ChannelId = c.Id AND m.SentAt > COALESCE((SELECT MAX(cm2.JoinedAt) FROM ChannelMembers cm2 WHERE cm2.ChannelId = c.Id AND cm2.UserId = @UserId), '1900-01-01') AND m.SenderId <> @UserId AND m.IsRead = 0) AS UnreadCount
                    FROM Channels c
                    LEFT JOIN Departments d ON c.DepartmentId = d.deptid
                    LEFT JOIN Users u ON c.CreatedBy = u.UserName
                    WHERE c.IsActive = 1 AND (c.IsPrivate = 0 OR EXISTS (SELECT 1 FROM ChannelMembers cm WHERE cm.ChannelId = c.Id AND cm.UserId = @UserId))
                    ORDER BY c.Name";

                var channels = (await db.QueryAsync<Channel>(sql, new { UserId = userId })).ToList();
                return channels;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting channels for user {UserId}", userId);
                return new List<Channel>();
            }
        }

        public async Task<Channel?> GetChannelByIdAsync(int channelId)
        {
            try
            {
                using var db = CreateConnection();
                var channelSql = @"
                    SELECT 
                        c.Id, c.Name, c.Description, c.DepartmentId, d.name AS DepartmentName,
                        c.IsPrivate, c.IsActive, c.CreatedBy, u.FullUserName AS CreatorName, c.CreatedAt
                    FROM Channels c
                    LEFT JOIN Departments d ON c.DepartmentId = d.deptid
                    LEFT JOIN Users u ON c.CreatedBy = u.UserName
                    WHERE c.Id = @ChannelId";

                var channel = await db.QueryFirstOrDefaultAsync<Channel>(channelSql, new { ChannelId = channelId });
                if (channel == null) return null;

                var membersSql = @"
                    SELECT 
                        cm.Id, cm.ChannelId, cm.UserId, cm.IsAdmin, cm.JoinedAt,
                        u.UserName, u.FullUserName AS FullName, e.Designation
                    FROM ChannelMembers cm
                    LEFT JOIN Users u ON cm.UserId = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    WHERE cm.ChannelId = @ChannelId
                    ORDER BY cm.IsAdmin DESC, u.UserName";

                channel.Members = (await db.QueryAsync<ChannelMember>(membersSql, new { ChannelId = channelId })).ToList();
                channel.MemberCount = channel.Members.Count;
                return channel;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting channel by ID {ChannelId}", channelId);
                return null;
            }
        }

        public async Task<int> CreateChannelAsync(Channel channel, List<string>? memberUserIds = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO Channels (Name, Description, DepartmentId, IsPrivate, IsActive, CreatedBy, CreatedAt)
                    VALUES (@Name, @Description, @DepartmentId, @IsPrivate, @IsActive, @CreatedBy, GETUTCDATE());
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                var channelId = await db.ExecuteScalarAsync<int>(sql, channel);
                
                // Add Creator as Admin
                await AddMemberToChannelAsync(channelId, channel.CreatedBy, true);

                // Add other members
                if (memberUserIds != null)
                {
                    foreach (var memberId in memberUserIds.Where(m => !string.Equals(m, channel.CreatedBy, StringComparison.OrdinalIgnoreCase)))
                    {
                        await AddMemberToChannelAsync(channelId, memberId, false);
                    }
                }

                return channelId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating channel {ChannelName}", channel.Name);
                return 0;
            }
        }

        public async Task<bool> AddMemberToChannelAsync(int channelId, string userId, bool isAdmin = false)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    IF NOT EXISTS (SELECT 1 FROM ChannelMembers WHERE ChannelId = @ChannelId AND UserId = @UserId)
                    BEGIN
                        INSERT INTO ChannelMembers (ChannelId, UserId, IsAdmin, JoinedAt)
                        VALUES (@ChannelId, @UserId, @IsAdmin, GETUTCDATE());
                    END";

                await db.ExecuteAsync(sql, new { ChannelId = channelId, UserId = userId, IsAdmin = isAdmin });
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding member {UserId} to channel {ChannelId}", userId, channelId);
                return false;
            }
        }

        public async Task<bool> RemoveMemberFromChannelAsync(int channelId, string userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "DELETE FROM ChannelMembers WHERE ChannelId = @ChannelId AND UserId = @UserId";
                await db.ExecuteAsync(sql, new { ChannelId = channelId, UserId = userId });
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error removing member {UserId} from channel {ChannelId}", userId, channelId);
                return false;
            }
        }

        #endregion

        #region 3. Messages & Direct Chat

        public async Task<List<Message>> GetChannelMessagesAsync(int channelId, int take = 50, long? beforeMessageId = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT TOP (@Take)
                        m.Id, m.ChannelId, m.SenderId, u.FullUserName AS SenderName, e.Designation AS SenderDesignation,
                        m.ReceiverId, m.Content, m.MessageType, m.IsRead, m.IsDeleted,
                        m.ParentMessageId, pm.Content AS ParentMessagePreview,
                        m.SentAt, m.EditedAt
                    FROM Messages m
                    LEFT JOIN Users u ON m.SenderId = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Messages pm ON m.ParentMessageId = pm.Id
                    WHERE m.ChannelId = @ChannelId AND (@BeforeId IS NULL OR m.Id < @BeforeId)
                    ORDER BY m.SentAt DESC";

                var messages = (await db.QueryAsync<Message>(sql, new { ChannelId = channelId, Take = take, BeforeId = beforeMessageId })).ToList();
                messages.Reverse(); // Display in chronological order

                if (messages.Any())
                {
                    var msgIds = messages.Select(m => m.Id).ToList();
                    var attachSql = "SELECT Id, MessageId, FileName, FilePath, FileSize, ContentType, UploadedAt FROM MessageAttachments WHERE MessageId IN @Ids";
                    var attachments = (await db.QueryAsync<MessageAttachment>(attachSql, new { Ids = msgIds })).ToList();

                    foreach (var msg in messages)
                    {
                        msg.Attachments = attachments.Where(a => a.MessageId == msg.Id).ToList();
                    }
                }

                return messages;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting channel messages for channel {ChannelId}", channelId);
                return new List<Message>();
            }
        }

        public async Task<List<Message>> GetDirectMessagesAsync(string userId1, string userId2, int take = 50, long? beforeMessageId = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT TOP (@Take)
                        m.Id, m.ChannelId, m.SenderId, u.FullUserName AS SenderName, e.Designation AS SenderDesignation,
                        m.ReceiverId, ru.FullUserName AS ReceiverName, m.Content, m.MessageType, m.IsRead, m.IsDeleted,
                        m.ParentMessageId, pm.Content AS ParentMessagePreview,
                        m.SentAt, m.EditedAt
                    FROM Messages m
                    LEFT JOIN Users u ON m.SenderId = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Users ru ON m.ReceiverId = ru.UserName
                    LEFT JOIN Messages pm ON m.ParentMessageId = pm.Id
                    WHERE m.ChannelId IS NULL 
                      AND ((m.SenderId = @User1 AND m.ReceiverId = @User2) OR (m.SenderId = @User2 AND m.ReceiverId = @User1))
                      AND (@BeforeId IS NULL OR m.Id < @BeforeId)
                    ORDER BY m.SentAt DESC";

                var messages = (await db.QueryAsync<Message>(sql, new { User1 = userId1, User2 = userId2, Take = take, BeforeId = beforeMessageId })).ToList();
                messages.Reverse();

                if (messages.Any())
                {
                    var msgIds = messages.Select(m => m.Id).ToList();
                    var attachSql = "SELECT Id, MessageId, FileName, FilePath, FileSize, ContentType, UploadedAt FROM MessageAttachments WHERE MessageId IN @Ids";
                    var attachments = (await db.QueryAsync<MessageAttachment>(attachSql, new { Ids = msgIds })).ToList();

                    foreach (var msg in messages)
                    {
                        msg.Attachments = attachments.Where(a => a.MessageId == msg.Id).ToList();
                    }
                }

                return messages;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting direct messages between {User1} and {User2}", userId1, userId2);
                return new List<Message>();
            }
        }

        public async Task<List<DirectChatSummary>> GetDirectChatSummariesAsync(string currentUserId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    WITH RankedMessages AS (
                        SELECT 
                            CASE WHEN SenderId = @UserId THEN ReceiverId ELSE SenderId END AS OtherUserId,
                            Content AS LastMessage,
                            SentAt AS LastMessageTime,
                            ROW_NUMBER() OVER (PARTITION BY CASE WHEN SenderId = @UserId THEN ReceiverId ELSE SenderId END ORDER BY SentAt DESC) AS rn
                        FROM Messages
                        WHERE ChannelId IS NULL AND (SenderId = @UserId OR ReceiverId = @UserId)
                    )
                    SELECT 
                        rm.OtherUserId,
                        u.UserName AS OtherUserName,
                        u.FullUserName AS OtherUserFullName,
                        e.Designation,
                        d.name AS DepartmentName,
                        COALESCE(p.Status, 0) AS Status,
                        rm.LastMessage,
                        rm.LastMessageTime,
                        (SELECT COUNT(1) FROM Messages m2 WHERE m2.ChannelId IS NULL AND m2.SenderId = rm.OtherUserId AND m2.ReceiverId = @UserId AND m2.IsRead = 0) AS UnreadCount
                    FROM RankedMessages rm
                    INNER JOIN Users u ON rm.OtherUserId = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    LEFT JOIN UserPresences p ON u.UserName = p.UserId
                    WHERE rm.rn = 1
                    ORDER BY rm.LastMessageTime DESC";

                var summaries = (await db.QueryAsync<DirectChatSummary>(sql, new { UserId = currentUserId })).ToList();
                return summaries;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting direct chat summaries for {UserId}", currentUserId);
                return new List<DirectChatSummary>();
            }
        }

        public async Task<long> SendMessageAsync(Message message)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO Messages (ChannelId, SenderId, ReceiverId, Content, MessageType, IsRead, IsDeleted, ParentMessageId, SentAt)
                    VALUES (@ChannelId, @SenderId, @ReceiverId, @Content, @MessageType, 0, 0, @ParentMessageId, GETUTCDATE());
                    SELECT CAST(SCOPE_IDENTITY() AS BIGINT);";

                var messageId = await db.ExecuteScalarAsync<long>(sql, message);

                if (message.Attachments != null && message.Attachments.Any())
                {
                    var attachSql = @"
                        INSERT INTO MessageAttachments (MessageId, FileName, FilePath, FileSize, ContentType, UploadedAt)
                        VALUES (@MessageId, @FileName, @FilePath, @FileSize, @ContentType, GETUTCDATE());";

                    foreach (var attach in message.Attachments)
                    {
                        attach.MessageId = messageId;
                        await db.ExecuteAsync(attachSql, attach);
                    }
                }

                return messageId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error sending message from {SenderId}", message.SenderId);
                return 0;
            }
        }

        public async Task<bool> MarkMessagesAsReadAsync(string currentUserId, string? senderId = null, int? channelId = null)
        {
            try
            {
                using var db = CreateConnection();
                if (channelId.HasValue)
                {
                    var sql = "UPDATE Messages SET IsRead = 1 WHERE ChannelId = @ChannelId AND SenderId <> @UserId AND IsRead = 0";
                    await db.ExecuteAsync(sql, new { ChannelId = channelId.Value, UserId = currentUserId });
                }
                else if (!string.IsNullOrEmpty(senderId))
                {
                    var sql = "UPDATE Messages SET IsRead = 1 WHERE ChannelId IS NULL AND SenderId = @SenderId AND ReceiverId = @UserId AND IsRead = 0";
                    await db.ExecuteAsync(sql, new { SenderId = senderId, UserId = currentUserId });
                }
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error marking messages read for {UserId}", currentUserId);
                return false;
            }
        }

        public async Task<bool> DeleteMessageAsync(long messageId, string userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "UPDATE Messages SET IsDeleted = 1, Content = 'This message was deleted' WHERE Id = @Id AND SenderId = @UserId";
                var rows = await db.ExecuteAsync(sql, new { Id = messageId, UserId = userId });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting message {MessageId}", messageId);
                return false;
            }
        }

        #endregion

        #region 4. Announcements

        public async Task<List<Announcement>> GetAnnouncementsAsync(string currentUserId, string? departmentId = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        a.Id, a.Title, a.Content, a.Priority, a.CreatedBy, u.FullUserName AS CreatorName,
                        e.Designation AS CreatorDesignation, a.DepartmentId, d.name AS DepartmentName,
                        a.IsPinned, a.IsActive, a.CreatedAt, a.UpdatedAt, a.ExpiresAt,
                        (SELECT COUNT(1) FROM AnnouncementAcknowledgments ack WHERE ack.AnnouncementId = a.Id) AS AcknowledgmentCount,
                        CASE WHEN EXISTS (SELECT 1 FROM AnnouncementAcknowledgments ack2 WHERE ack2.AnnouncementId = a.Id AND ack2.UserId = @UserId) THEN 1 ELSE 0 END AS IsAcknowledgedByCurrentUser
                    FROM Announcements a
                    LEFT JOIN Users u ON a.CreatedBy = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON a.DepartmentId = d.deptid
                    WHERE a.IsActive = 1 AND (@DeptId IS NULL OR a.DepartmentId IS NULL OR a.DepartmentId = @DeptId)
                    ORDER BY a.IsPinned DESC, a.CreatedAt DESC";

                var announcements = (await db.QueryAsync<Announcement>(sql, new { UserId = currentUserId, DeptId = departmentId })).ToList();

                if (announcements.Any())
                {
                    var aIds = announcements.Select(a => a.Id).ToList();
                    var attachSql = "SELECT Id, AnnouncementId, FileName, FilePath, FileSize, ContentType, UploadedAt FROM AnnouncementAttachments WHERE AnnouncementId IN @Ids";
                    var attachments = (await db.QueryAsync<AnnouncementAttachment>(attachSql, new { Ids = aIds })).ToList();

                    foreach (var a in announcements)
                    {
                        a.Attachments = attachments.Where(att => att.AnnouncementId == a.Id).ToList();
                    }
                }

                return announcements;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting announcements");
                return new List<Announcement>();
            }
        }

        public async Task<Announcement?> GetAnnouncementByIdAsync(int id, string currentUserId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        a.Id, a.Title, a.Content, a.Priority, a.CreatedBy, u.FullUserName AS CreatorName,
                        e.Designation AS CreatorDesignation, a.DepartmentId, d.name AS DepartmentName,
                        a.IsPinned, a.IsActive, a.CreatedAt, a.UpdatedAt, a.ExpiresAt,
                        (SELECT COUNT(1) FROM AnnouncementAcknowledgments ack WHERE ack.AnnouncementId = a.Id) AS AcknowledgmentCount,
                        CASE WHEN EXISTS (SELECT 1 FROM AnnouncementAcknowledgments ack2 WHERE ack2.AnnouncementId = a.Id AND ack2.UserId = @UserId) THEN 1 ELSE 0 END AS IsAcknowledgedByCurrentUser
                    FROM Announcements a
                    LEFT JOIN Users u ON a.CreatedBy = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON a.DepartmentId = d.deptid
                    WHERE a.Id = @Id";

                var announcement = await db.QueryFirstOrDefaultAsync<Announcement>(sql, new { Id = id, UserId = currentUserId });
                if (announcement == null) return null;

                var attachSql = "SELECT Id, AnnouncementId, FileName, FilePath, FileSize, ContentType, UploadedAt FROM AnnouncementAttachments WHERE AnnouncementId = @Id";
                announcement.Attachments = (await db.QueryAsync<AnnouncementAttachment>(attachSql, new { Id = id })).ToList();

                var ackSql = @"
                    SELECT ack.Id, ack.AnnouncementId, ack.UserId, u.UserName, u.FullUserName AS FullName, ack.AcknowledgedAt
                    FROM AnnouncementAcknowledgments ack
                    LEFT JOIN Users u ON ack.UserId = u.UserName
                    WHERE ack.AnnouncementId = @Id
                    ORDER BY ack.AcknowledgedAt DESC";
                announcement.Acknowledgments = (await db.QueryAsync<AnnouncementAcknowledgment>(ackSql, new { Id = id })).ToList();

                return announcement;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting announcement {Id}", id);
                return null;
            }
        }

        public async Task<int> CreateAnnouncementAsync(Announcement announcement)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO Announcements (Title, Content, Priority, CreatedBy, DepartmentId, IsPinned, IsActive, CreatedAt, ExpiresAt)
                    VALUES (@Title, @Content, @Priority, @CreatedBy, @DepartmentId, @IsPinned, 1, GETUTCDATE(), @ExpiresAt);
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                var aId = await db.ExecuteScalarAsync<int>(sql, announcement);

                if (announcement.Attachments != null && announcement.Attachments.Any())
                {
                    var attachSql = @"
                        INSERT INTO AnnouncementAttachments (AnnouncementId, FileName, FilePath, FileSize, ContentType, UploadedAt)
                        VALUES (@AnnouncementId, @FileName, @FilePath, @FileSize, @ContentType, GETUTCDATE());";

                    foreach (var att in announcement.Attachments)
                    {
                        att.AnnouncementId = aId;
                        await db.ExecuteAsync(attachSql, att);
                    }
                }

                return aId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating announcement");
                return 0;
            }
        }

        public async Task<bool> AcknowledgeAnnouncementAsync(int announcementId, string userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    IF NOT EXISTS (SELECT 1 FROM AnnouncementAcknowledgments WHERE AnnouncementId = @AnnouncementId AND UserId = @UserId)
                    BEGIN
                        INSERT INTO AnnouncementAcknowledgments (AnnouncementId, UserId, AcknowledgedAt)
                        VALUES (@AnnouncementId, @UserId, GETUTCDATE());
                    END";

                await db.ExecuteAsync(sql, new { AnnouncementId = announcementId, UserId = userId });
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error acknowledging announcement {AnnouncementId} by {UserId}", announcementId, userId);
                return false;
            }
        }

        #endregion

        #region 5. Tasks

        public async Task<List<TaskItem>> GetTasksAsync(string? assignedTo = null, string? assignedBy = null, string? departmentId = null, TaskItemStatus? status = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        t.Id, t.Title, t.Description, t.AssignedTo, u.FullUserName AS AssigneeName, e.Designation AS AssigneeDesignation,
                        t.AssignedBy, bu.FullUserName AS AssignerName, t.DepartmentId, d.name AS DepartmentName,
                        t.Priority, t.Status, t.DueDate, t.WhatsAppMessageSent, t.CreatedAt, t.UpdatedAt, t.CompletedAt
                    FROM TaskItems t
                    LEFT JOIN Users u ON t.AssignedTo = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Users bu ON t.AssignedBy = bu.UserName
                    LEFT JOIN Departments d ON t.DepartmentId = d.deptid
                    WHERE (@AssignedTo IS NULL OR t.AssignedTo = @AssignedTo)
                      AND (@AssignedBy IS NULL OR t.AssignedBy = @AssignedBy)
                      AND (@DeptId IS NULL OR t.DepartmentId = @DeptId)
                      AND (@Status IS NULL OR t.Status = @Status)
                    ORDER BY t.Priority DESC, t.DueDate ASC, t.CreatedAt DESC";

                var tasks = (await db.QueryAsync<TaskItem>(sql, new { AssignedTo = assignedTo, AssignedBy = assignedBy, DeptId = departmentId, Status = (int?)status })).ToList();
                return tasks;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting tasks");
                return new List<TaskItem>();
            }
        }

        public async Task<TaskItem?> GetTaskByIdAsync(int taskId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        t.Id, t.Title, t.Description, t.AssignedTo, u.FullUserName AS AssigneeName, e.Designation AS AssigneeDesignation,
                        t.AssignedBy, bu.FullUserName AS AssignerName, t.DepartmentId, d.name AS DepartmentName,
                        t.Priority, t.Status, t.DueDate, t.WhatsAppMessageSent, t.CreatedAt, t.UpdatedAt, t.CompletedAt
                    FROM TaskItems t
                    LEFT JOIN Users u ON t.AssignedTo = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Users bu ON t.AssignedBy = bu.UserName
                    LEFT JOIN Departments d ON t.DepartmentId = d.deptid
                    WHERE t.Id = @Id";

                var task = await db.QueryFirstOrDefaultAsync<TaskItem>(sql, new { Id = taskId });
                if (task == null) return null;

                var commentSql = @"
                    SELECT tc.Id, tc.TaskId, tc.UserId, u.UserName, u.FullUserName, tc.Content, tc.CreatedAt
                    FROM TaskComments tc
                    LEFT JOIN Users u ON tc.UserId = u.UserName
                    WHERE tc.TaskId = @TaskId
                    ORDER BY tc.CreatedAt ASC";
                task.Comments = (await db.QueryAsync<TaskComment>(commentSql, new { TaskId = taskId })).ToList();

                var attachSql = "SELECT Id, TaskId, FileName, FilePath, FileSize, ContentType, UploadedAt FROM TaskAttachments WHERE TaskId = @TaskId";
                task.Attachments = (await db.QueryAsync<TaskAttachment>(attachSql, new { TaskId = taskId })).ToList();

                return task;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting task {TaskId}", taskId);
                return null;
            }
        }

        public async Task<int> CreateTaskAsync(TaskItem task)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO TaskItems (Title, Description, AssignedTo, AssignedBy, DepartmentId, Priority, Status, DueDate, WhatsAppMessageSent, CreatedAt)
                    VALUES (@Title, @Description, @AssignedTo, @AssignedBy, @DepartmentId, @Priority, @Status, @DueDate, @WhatsAppMessageSent, GETUTCDATE());
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                var taskId = await db.ExecuteScalarAsync<int>(sql, task);

                if (task.Attachments != null && task.Attachments.Any())
                {
                    var attachSql = @"
                        INSERT INTO TaskAttachments (TaskId, FileName, FilePath, FileSize, ContentType, UploadedAt)
                        VALUES (@TaskId, @FileName, @FilePath, @FileSize, @ContentType, GETUTCDATE());";

                    foreach (var att in task.Attachments)
                    {
                        att.TaskId = taskId;
                        await db.ExecuteAsync(attachSql, att);
                    }
                }

                return taskId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating task {Title}", task.Title);
                return 0;
            }
        }

        public async Task<bool> UpdateTaskStatusAsync(int taskId, TaskItemStatus status, DateTime? completedAt = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    UPDATE TaskItems 
                    SET Status = @Status, 
                        CompletedAt = CASE WHEN @Status = 2 THEN COALESCE(@CompletedAt, GETUTCDATE()) ELSE NULL END,
                        UpdatedAt = GETUTCDATE()
                    WHERE Id = @Id";

                var rows = await db.ExecuteAsync(sql, new { Id = taskId, Status = (int)status, CompletedAt = completedAt });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating status for task {TaskId}", taskId);
                return false;
            }
        }

        public async Task<int> AddTaskCommentAsync(TaskComment comment)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                    VALUES (@TaskId, @UserId, @Content, GETUTCDATE());
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                return await db.ExecuteScalarAsync<int>(sql, comment);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding comment to task {TaskId}", comment.TaskId);
                return 0;
            }
        }

        #endregion

        #region 6. Meetings

        public async Task<List<Meeting>> GetMeetingsForUserAsync(string userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        m.Id, m.Title, m.OrganizerId, u.FullUserName AS OrganizerName,
                        m.MeetingUrl, m.ScheduledStartTime, m.ScheduledEndTime, m.Status,
                        m.MeetingMinutes, m.IsReminderSent, m.CreatedAt
                    FROM Meetings m
                    LEFT JOIN Users u ON m.OrganizerId = u.UserName
                    WHERE m.OrganizerId = @UserId 
                       OR EXISTS (SELECT 1 FROM MeetingParticipants mp WHERE mp.MeetingId = m.Id AND mp.UserId = @UserId)
                    ORDER BY m.ScheduledStartTime DESC";

                var meetings = (await db.QueryAsync<Meeting>(sql, new { UserId = userId })).ToList();
                return meetings;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting meetings for user {UserId}", userId);
                return new List<Meeting>();
            }
        }

        public async Task<Meeting?> GetMeetingByIdAsync(int meetingId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        m.Id, m.Title, m.OrganizerId, u.FullUserName AS OrganizerName,
                        m.MeetingUrl, m.ScheduledStartTime, m.ScheduledEndTime, m.Status,
                        m.MeetingMinutes, m.IsReminderSent, m.CreatedAt
                    FROM Meetings m
                    LEFT JOIN Users u ON m.OrganizerId = u.UserName
                    WHERE m.Id = @Id";

                var meeting = await db.QueryFirstOrDefaultAsync<Meeting>(sql, new { Id = meetingId });
                if (meeting == null) return null;

                var partSql = @"
                    SELECT mp.MeetingId, mp.UserId, u.UserName, u.FullUserName AS FullName, e.Designation, mp.HasAttended
                    FROM MeetingParticipants mp
                    LEFT JOIN Users u ON mp.UserId = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    WHERE mp.MeetingId = @MeetingId";

                meeting.Participants = (await db.QueryAsync<MeetingParticipant>(partSql, new { MeetingId = meetingId })).ToList();
                return meeting;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting meeting {MeetingId}", meetingId);
                return null;
            }
        }

        public async Task<int> CreateMeetingAsync(Meeting meeting, List<string>? participantUserIds = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO Meetings (Title, OrganizerId, MeetingUrl, ScheduledStartTime, ScheduledEndTime, Status, MeetingMinutes, IsReminderSent, CreatedAt)
                    VALUES (@Title, @OrganizerId, @MeetingUrl, @ScheduledStartTime, @ScheduledEndTime, @Status, @MeetingMinutes, 0, GETUTCDATE());
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                var meetingId = await db.ExecuteScalarAsync<int>(sql, meeting);

                // Add organizer as participant
                var addPartSql = "INSERT INTO MeetingParticipants (MeetingId, UserId, HasAttended) VALUES (@MeetingId, @UserId, 0)";
                await db.ExecuteAsync(addPartSql, new { MeetingId = meetingId, UserId = meeting.OrganizerId });

                if (participantUserIds != null)
                {
                    foreach (var pId in participantUserIds.Where(p => !string.Equals(p, meeting.OrganizerId, StringComparison.OrdinalIgnoreCase)))
                    {
                        await db.ExecuteAsync(addPartSql, new { MeetingId = meetingId, UserId = pId });
                    }
                }

                return meetingId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating meeting {Title}", meeting.Title);
                return 0;
            }
        }

        public async Task<bool> UpdateMeetingStatusAsync(int meetingId, MeetingStatus status)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "UPDATE Meetings SET Status = @Status WHERE Id = @Id";
                var rows = await db.ExecuteAsync(sql, new { Id = meetingId, Status = (int)status });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating meeting status for {MeetingId}", meetingId);
                return false;
            }
        }

        #endregion

        #region 7. Minutes Approval & Workflow

        public async Task<List<MinuteApproval>> GetMinutesAsync(MinuteFilter filter)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        m.Id, m.Date, m.No, m.Type, m.Subject, m.Points, m.ForwardToUserId,
                        fu.FullUserName AS ForwardToUserName, m.Currency, m.TotalAmount,
                        m.AdvancePercentage, m.AdvanceAmount, m.IsUrgent, m.CloseByInitiator,
                        m.CreatedByUserId, cu.FullUserName AS CreatedByUserName, m.CreatedAt,
                        m.Status, m.SignaturePath, m.PurchaseOrderValue, m.PurchaseAdvanceRecommend,
                        m.PurchaseApprovedAmount, m.RequestedStockQty, m.CurrentStockQty,
                        m.ApprovedStockQty, m.HRLeaveType, m.FinancialType
                    FROM MinuteApprovals m
                    LEFT JOIN Users cu ON m.CreatedByUserId = cu.UserName
                    LEFT JOIN Users fu ON m.ForwardToUserId = fu.UserName
                    WHERE (@FromDate IS NULL OR m.Date >= @FromDate)
                      AND (@ToDate IS NULL OR m.Date <= @ToDate)
                      AND (@No IS NULL OR m.No LIKE @NoParam)
                      AND (@Type IS NULL OR m.Type = @Type)
                      AND (@Status IS NULL OR m.Status = @Status)
                      AND (@UserId IS NULL OR m.CreatedByUserId = @UserId OR m.ForwardToUserId = @UserId)
                      AND (@SearchTerm IS NULL OR m.Subject LIKE @SearchParam OR m.Points LIKE @SearchParam)
                    ORDER BY m.Date DESC, m.CreatedAt DESC";

                var noParam = string.IsNullOrWhiteSpace(filter.No) ? null : $"%{filter.No.Trim()}%";
                var searchParam = string.IsNullOrWhiteSpace(filter.SearchTerm) ? null : $"%{filter.SearchTerm.Trim()}%";

                var minutes = (await db.QueryAsync<MinuteApproval>(sql, new
                {
                    filter.FromDate,
                    filter.ToDate,
                    filter.No,
                    NoParam = noParam,
                    filter.Type,
                    filter.Status,
                    filter.UserId,
                    filter.SearchTerm,
                    SearchParam = searchParam
                })).ToList();

                return minutes;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting minutes");
                return new List<MinuteApproval>();
            }
        }

        public async Task<MinuteApproval?> GetMinuteByIdAsync(int id)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        m.Id, m.Date, m.No, m.Type, m.Subject, m.Points, m.ForwardToUserId,
                        fu.FullUserName AS ForwardToUserName, m.Currency, m.TotalAmount,
                        m.AdvancePercentage, m.AdvanceAmount, m.IsUrgent, m.CloseByInitiator,
                        m.CreatedByUserId, cu.FullUserName AS CreatedByUserName, m.CreatedAt,
                        m.Status, m.SignaturePath, m.PurchaseOrderValue, m.PurchaseAdvanceRecommend,
                        m.PurchaseApprovedAmount, m.RequestedStockQty, m.CurrentStockQty,
                        m.ApprovedStockQty, m.HRLeaveType, m.FinancialType
                    FROM MinuteApprovals m
                    LEFT JOIN Users cu ON m.CreatedByUserId = cu.UserName
                    LEFT JOIN Users fu ON m.ForwardToUserId = fu.UserName
                    WHERE m.Id = @Id";

                var minute = await db.QueryFirstOrDefaultAsync<MinuteApproval>(sql, new { Id = id });
                if (minute == null) return null;

                var attachSql = "SELECT Id, MinuteApprovalId, FileName, FilePath, FileSize, ContentType, UploadedAt FROM MinuteAttachments WHERE MinuteApprovalId = @Id";
                minute.Attachments = (await db.QueryAsync<MinuteAttachment>(attachSql, new { Id = id })).ToList();

                var histSql = @"
                    SELECT h.Id, h.MinuteApprovalId, h.UserId, u.UserName, u.FullUserName AS FullName, e.Designation, h.ActionTaken, h.Remarks, h.SignaturePath, h.CreatedAt
                    FROM MinuteWorkflowHistories h
                    LEFT JOIN Users u ON h.UserId = u.UserName
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    WHERE h.MinuteApprovalId = @Id
                    ORDER BY h.CreatedAt ASC";
                minute.WorkflowHistory = (await db.QueryAsync<MinuteWorkflowHistory>(histSql, new { Id = id })).ToList();

                return minute;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting minute {Id}", id);
                return null;
            }
        }

        public async Task<int> CreateMinuteApprovalAsync(MinuteApproval minute)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO MinuteApprovals (Date, No, Type, Subject, Points, ForwardToUserId, Currency, TotalAmount, AdvancePercentage, AdvanceAmount, IsUrgent, CloseByInitiator, CreatedByUserId, CreatedAt, Status, SignaturePath, PurchaseOrderValue, PurchaseAdvanceRecommend, PurchaseApprovedAmount, RequestedStockQty, CurrentStockQty, ApprovedStockQty, HRLeaveType, FinancialType)
                    VALUES (@Date, @No, @Type, @Subject, @Points, @ForwardToUserId, @Currency, @TotalAmount, @AdvancePercentage, @AdvanceAmount, @IsUrgent, @CloseByInitiator, @CreatedByUserId, GETUTCDATE(), @Status, @SignaturePath, @PurchaseOrderValue, @PurchaseAdvanceRecommend, @PurchaseApprovedAmount, @RequestedStockQty, @CurrentStockQty, @ApprovedStockQty, @HRLeaveType, @FinancialType);
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                var minuteId = await db.ExecuteScalarAsync<int>(sql, minute);

                if (minute.Attachments != null && minute.Attachments.Any())
                {
                    var attachSql = @"
                        INSERT INTO MinuteAttachments (MinuteApprovalId, FileName, FilePath, FileSize, ContentType, UploadedAt)
                        VALUES (@MinuteApprovalId, @FileName, @FilePath, @FileSize, @ContentType, GETUTCDATE());";

                    foreach (var att in minute.Attachments)
                    {
                        att.MinuteApprovalId = minuteId;
                        await db.ExecuteAsync(attachSql, att);
                    }
                }

                // Add initial workflow history
                var histSql = @"
                    INSERT INTO MinuteWorkflowHistories (MinuteApprovalId, UserId, ActionTaken, Remarks, SignaturePath, CreatedAt)
                    VALUES (@MinuteApprovalId, @UserId, 'Created', 'Minute created and initiated', @SignaturePath, GETUTCDATE());";
                await db.ExecuteAsync(histSql, new { MinuteApprovalId = minuteId, UserId = minute.CreatedByUserId, SignaturePath = minute.SignaturePath });

                return minuteId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating minute approval");
                return 0;
            }
        }

        public async Task<bool> UpdateMinuteStatusAsync(int minuteId, string status, string userId, string actionTaken, string? remarks = null, string? signaturePath = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "UPDATE MinuteApprovals SET Status = @Status WHERE Id = @Id";
                await db.ExecuteAsync(sql, new { Id = minuteId, Status = status });

                var histSql = @"
                    INSERT INTO MinuteWorkflowHistories (MinuteApprovalId, UserId, ActionTaken, Remarks, SignaturePath, CreatedAt)
                    VALUES (@MinuteApprovalId, @UserId, @ActionTaken, @Remarks, @SignaturePath, GETUTCDATE());";

                await db.ExecuteAsync(histSql, new
                {
                    MinuteApprovalId = minuteId,
                    UserId = userId,
                    ActionTaken = actionTaken,
                    Remarks = remarks,
                    SignaturePath = signaturePath
                });

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating status for minute {MinuteId}", minuteId);
                return false;
            }
        }

        public async Task<List<MinuteType>> GetMinuteTypesAsync()
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT Id, Name, IsActive, CreatedAt FROM MinuteTypes WHERE IsActive = 1 ORDER BY Name";
                return (await db.QueryAsync<MinuteType>(sql)).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting minute types");
                return new List<MinuteType>();
            }
        }

        #endregion

        #region 8. Presence & Sticky Notes

        public async Task<bool> UpdateUserPresenceAsync(string userId, UserStatus status, string? connectionId = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    IF EXISTS (SELECT 1 FROM UserPresences WHERE UserId = @UserId)
                    BEGIN
                        UPDATE UserPresences 
                        SET Status = @Status, 
                            LastSeen = GETUTCDATE(), 
                            ConnectionId = COALESCE(@ConnectionId, ConnectionId)
                        WHERE UserId = @UserId;
                    END
                    ELSE
                    BEGIN
                        INSERT INTO UserPresences (UserId, Status, LastSeen, ConnectionId)
                        VALUES (@UserId, @Status, GETUTCDATE(), @ConnectionId);
                    END";

                await db.ExecuteAsync(sql, new { UserId = userId, Status = (int)status, ConnectionId = connectionId });
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating presence for {UserId}", userId);
                return false;
            }
        }

        public async Task<List<UserPresence>> GetUserPresencesAsync()
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT UserId, Status, LastSeen, ConnectionId FROM UserPresences";
                return (await db.QueryAsync<UserPresence>(sql)).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting user presences");
                return new List<UserPresence>();
            }
        }

        public async Task<List<StickyNote>> GetStickyNotesAsync(string userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT Id, UserId, Content, Color, XPos, YPos, ReminderTime, IsReminderSent, CreatedAt FROM StickyNotes WHERE UserId = @UserId ORDER BY CreatedAt DESC";
                return (await db.QueryAsync<StickyNote>(sql, new { UserId = userId })).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting sticky notes for {UserId}", userId);
                return new List<StickyNote>();
            }
        }

        public async Task<int> SaveStickyNoteAsync(StickyNote note)
        {
            try
            {
                using var db = CreateConnection();
                if (note.Id == 0)
                {
                    var sql = @"
                        INSERT INTO StickyNotes (UserId, Content, Color, XPos, YPos, ReminderTime, IsReminderSent, CreatedAt)
                        VALUES (@UserId, @Content, @Color, @XPos, @YPos, @ReminderTime, @IsReminderSent, GETUTCDATE());
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";
                    return await db.ExecuteScalarAsync<int>(sql, note);
                }
                else
                {
                    var sql = @"
                        UPDATE StickyNotes 
                        SET Content = @Content, Color = @Color, XPos = @XPos, YPos = @YPos, ReminderTime = @ReminderTime
                        WHERE Id = @Id AND UserId = @UserId";
                    await db.ExecuteAsync(sql, note);
                    return note.Id;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving sticky note for {UserId}", note.UserId);
                return 0;
            }
        }

        public async Task<bool> DeleteStickyNoteAsync(int id, string userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "DELETE FROM StickyNotes WHERE Id = @Id AND UserId = @UserId";
                var rows = await db.ExecuteAsync(sql, new { Id = id, UserId = userId });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting sticky note {Id}", id);
                return false;
            }
        }

        #endregion

        #region 9. Leads & Pipeline

        public async Task<List<LeadModel>> GetLeadsAsync(string? status = null, string? priority = null, string? search = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT Id, LeadNumber, CompanyName, ContactPerson, Designation, Email, Phone, WhatsApp,
                           Country, City, Website, Source, Industry, ProductInterest, EstimatedQuantity,
                           EstimatedValue, Currency, ExpectedOrderDate, Status, Priority, Notes,
                           NextFollowUpDate, ConvertedCustCode, AssignedTo, CreatedAt, UpdatedAt
                    FROM Leads
                    WHERE (@Status IS NULL OR @Status = 'All' OR Status = @Status)
                      AND (@Priority IS NULL OR @Priority = 'All' OR Priority = @Priority)
                      AND (@Search IS NULL OR CompanyName LIKE @SearchParam OR ContactPerson LIKE @SearchParam OR Country LIKE @SearchParam OR LeadNumber LIKE @SearchParam)
                    ORDER BY UpdatedAt DESC";

                var searchParam = string.IsNullOrWhiteSpace(search) ? null : $"%{search.Trim()}%";
                var leads = (await db.QueryAsync<LeadModel>(sql, new
                {
                    Status = status,
                    Priority = priority,
                    Search = search,
                    SearchParam = searchParam
                })).ToList();

                return leads;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting leads");
                return new List<LeadModel>();
            }
        }

        public async Task<LeadModel?> GetLeadByIdAsync(int id)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"SELECT * FROM Leads WHERE Id = @Id";
                var lead = await db.QueryFirstOrDefaultAsync<LeadModel>(sql, new { Id = id });
                if (lead != null)
                {
                    lead.Activities = await GetLeadActivitiesAsync(id);
                }
                return lead;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting lead {Id}", id);
                return null;
            }
        }

        public async Task<int> SaveLeadAsync(LeadModel lead)
        {
            try
            {
                using var db = CreateConnection();
                if (lead.Id == 0)
                {
                    if (string.IsNullOrWhiteSpace(lead.LeadNumber))
                    {
                        lead.LeadNumber = $"LD-{DateTime.UtcNow:yyyyMMdd}-{Guid.NewGuid().ToString()[..4].ToUpper()}";
                    }

                    var insertSql = @"
                        INSERT INTO Leads (
                            LeadNumber, CompanyName, ContactPerson, Designation, Email, Phone, WhatsApp,
                            Country, City, Website, Source, Industry, ProductInterest, EstimatedQuantity,
                            EstimatedValue, Currency, ExpectedOrderDate, Status, Priority, Notes,
                            NextFollowUpDate, ConvertedCustCode, AssignedTo, CreatedAt, UpdatedAt
                        ) VALUES (
                            @LeadNumber, @CompanyName, @ContactPerson, @Designation, @Email, @Phone, @WhatsApp,
                            @Country, @City, @Website, @Source, @Industry, @ProductInterest, @EstimatedQuantity,
                            @EstimatedValue, @Currency, @ExpectedOrderDate, @Status, @Priority, @Notes,
                            @NextFollowUpDate, @ConvertedCustCode, @AssignedTo, GETUTCDATE(), GETUTCDATE()
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    return await db.ExecuteScalarAsync<int>(insertSql, lead);
                }
                else
                {
                    var updateSql = @"
                        UPDATE Leads
                        SET CompanyName = @CompanyName,
                            ContactPerson = @ContactPerson,
                            Designation = @Designation,
                            Email = @Email,
                            Phone = @Phone,
                            WhatsApp = @WhatsApp,
                            Country = @Country,
                            City = @City,
                            Website = @Website,
                            Source = @Source,
                            Industry = @Industry,
                            ProductInterest = @ProductInterest,
                            EstimatedQuantity = @EstimatedQuantity,
                            EstimatedValue = @EstimatedValue,
                            Currency = @Currency,
                            ExpectedOrderDate = @ExpectedOrderDate,
                            Status = @Status,
                            Priority = @Priority,
                            Notes = @Notes,
                            NextFollowUpDate = @NextFollowUpDate,
                            AssignedTo = @AssignedTo,
                            UpdatedAt = GETUTCDATE()
                        WHERE Id = @Id;";

                    await db.ExecuteAsync(updateSql, lead);
                    return lead.Id;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving lead");
                return 0;
            }
        }

        public async Task<bool> DeleteLeadAsync(int id)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "DELETE FROM Leads WHERE Id = @Id";
                var rows = await db.ExecuteAsync(sql, new { Id = id });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting lead {Id}", id);
                return false;
            }
        }

        public async Task<List<LeadActivityModel>> GetLeadActivitiesAsync(int leadId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT * FROM LeadActivities WHERE LeadId = @LeadId ORDER BY ActivityDate DESC";
                var list = await db.QueryAsync<LeadActivityModel>(sql, new { LeadId = leadId });
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting activities for lead {LeadId}", leadId);
                return new List<LeadActivityModel>();
            }
        }

        public async Task<int> AddLeadActivityAsync(LeadActivityModel activity)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO LeadActivities (LeadId, ActivityType, Description, PerformedBy, ActivityDate, NextFollowUpDate)
                    VALUES (@LeadId, @ActivityType, @Description, @PerformedBy, GETUTCDATE(), @NextFollowUpDate);
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                var id = await db.ExecuteScalarAsync<int>(sql, activity);

                if (activity.NextFollowUpDate.HasValue)
                {
                    await db.ExecuteAsync("UPDATE Leads SET NextFollowUpDate = @NextDate, UpdatedAt = GETUTCDATE() WHERE Id = @LeadId",
                        new { NextDate = activity.NextFollowUpDate.Value, LeadId = activity.LeadId });
                }

                return id;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding lead activity");
                return 0;
            }
        }

        public async Task<bool> CheckCustCodeExistsAsync(string custCode)
        {
            try
            {
                using var db = CreateConnection();
                var count = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(1) FROM ForeignCustomers WHERE CustCode = @CustCode",
                    new { CustCode = custCode.Trim() });
                return count > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking CustCode {CustCode}", custCode);
                return false;
            }
        }

        public async Task<bool> ConvertLeadToForeignCustomerAsync(ConvertLeadToCustomerModel model)
        {
            using var db = CreateConnection();
            db.Open();
            using var trans = db.BeginTransaction();
            try
            {
                // 1. Insert into ForeignCustomers
                var insertCustSql = @"
                    INSERT INTO ForeignCustomers (
                        CustCode, Country, Name, Address, City, 
                        Phone1, Email1, Cont1name, cont1Mobile, CustomerSource, Active
                    ) VALUES (
                        @CustCode, @Country, @Name, @Address, @City,
                        @Phone1, @Email1, @Cont1name, @Cont1Mobile, @CustomerSource, 1
                    );";

                await db.ExecuteAsync(insertCustSql, new
                {
                    CustCode = model.CustCode.Trim(),
                    Country = model.Country.Trim(),
                    Name = model.Name.Trim(),
                    Address = model.Address ?? "",
                    City = model.City ?? "",
                    Phone1 = model.Phone1 ?? "",
                    Email1 = model.Email1 ?? "",
                    Cont1name = model.Cont1name ?? "",
                    Cont1Mobile = model.Cont1Mobile ?? "",
                    CustomerSource = model.CustomerSource ?? "Lead Conversion"
                }, transaction: trans);

                // 2. Update Lead status to 'Converted'
                var updateLeadSql = @"
                    UPDATE Leads
                    SET Status = 'Converted', ConvertedCustCode = @CustCode, UpdatedAt = GETUTCDATE()
                    WHERE Id = @LeadId;";

                await db.ExecuteAsync(updateLeadSql, new
                {
                    LeadId = model.LeadId,
                    CustCode = model.CustCode.Trim()
                }, transaction: trans);

                // 3. Add activity log
                var activitySql = @"
                    INSERT INTO LeadActivities (LeadId, ActivityType, Description, PerformedBy, ActivityDate)
                    VALUES (@LeadId, 'Conversion', 'Converted to Foreign Customer [' + @CustCode + ']', 'System', GETUTCDATE());";

                await db.ExecuteAsync(activitySql, new
                {
                    LeadId = model.LeadId,
                    CustCode = model.CustCode.Trim()
                }, transaction: trans);

                trans.Commit();
                return true;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                _logger.LogError(ex, "Error converting lead {LeadId} to customer {CustCode}", model.LeadId, model.CustCode);
                throw;
            }
        }

        #endregion
    }
}
