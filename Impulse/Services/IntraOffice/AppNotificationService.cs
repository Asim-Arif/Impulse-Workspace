using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.IntraOffice
{
    public enum NotificationCategory
    {
        Message,
        Announcement,
        Minute,
        Meeting,
        Task,
        System,
        Lead
    }

    public class AppNotification
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public NotificationCategory Category { get; set; } = NotificationCategory.System;
        public string Title { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public string? SenderName { get; set; }
        public string? TargetUserId { get; set; } // Null for broadcast to all users
        public string ActionUrl { get; set; } = "/";
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public bool IsRead { get; set; } = false;
        /// <summary>True when this notification is a read-receipt (someone read your content)</summary>
        public bool IsReadReceipt { get; set; } = false;
    }

    public interface IAppNotificationService
    {
        event Func<AppNotification, Task>? OnNotificationReceived;
        Task SendNotificationAsync(AppNotification notification);
        List<AppNotification> GetRecentNotifications(string userId);
        int GetUnreadCount(string userId);
        void MarkAsRead(string notificationId, string userId);
        void MarkAllAsRead(string userId);
    }

    public class AppNotificationService : IAppNotificationService
    {
        public event Func<AppNotification, Task>? OnNotificationReceived;
        private readonly string _connectionString;
        private readonly ILogger<AppNotificationService> _logger;

        public AppNotificationService(IConfiguration configuration, ILogger<AppNotificationService> logger)
        {
            _logger = logger;
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task SendNotificationAsync(AppNotification notification)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    INSERT INTO AppNotifications (Id, Category, Title, Message, SenderName, TargetUserId, ActionUrl, CreatedAt, IsRead, IsReadReceipt)
                    VALUES (@Id, @Category, @Title, @Message, @SenderName, @TargetUserId, @ActionUrl, @CreatedAt, @IsRead, @IsReadReceipt);";

                await db.ExecuteAsync(sql, notification);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to persist AppNotification #{Id} to SQL database", notification.Id);
            }

            // Real-time dispatch to active client circuits
            if (OnNotificationReceived != null)
            {
                var delegates = OnNotificationReceived.GetInvocationList();
                foreach (Func<AppNotification, Task> handler in delegates)
                {
                    try
                    {
                        await handler(notification);
                    }
                    catch
                    {
                        // Ignore disconnected circuit errors
                    }
                }
            }
        }

        public List<AppNotification> GetRecentNotifications(string userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT TOP 30 
                        Id, 
                        Category, 
                        Title, 
                        Message, 
                        SenderName, 
                        TargetUserId, 
                        ActionUrl, 
                        CreatedAt, 
                        IsRead, 
                        IsReadReceipt
                    FROM AppNotifications
                    WHERE (TargetUserId IS NULL OR LOWER(LTRIM(RTRIM(TargetUserId))) = LOWER(LTRIM(RTRIM(@UserId))))
                    ORDER BY CreatedAt DESC;";

                return db.Query<AppNotification>(sql, new { UserId = userId?.Trim() }).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching recent notifications for user '{UserId}' from database", userId);
                return new List<AppNotification>();
            }
        }

        public int GetUnreadCount(string userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT COUNT(*) 
                    FROM AppNotifications
                    WHERE (TargetUserId IS NULL OR LOWER(LTRIM(RTRIM(TargetUserId))) = LOWER(LTRIM(RTRIM(@UserId))))
                      AND IsRead = 0;";

                return db.ExecuteScalar<int>(sql, new { UserId = userId?.Trim() });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error counting unread notifications for user '{UserId}' from database", userId);
                return 0;
            }
        }

        public void MarkAsRead(string notificationId, string userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    UPDATE AppNotifications 
                    SET IsRead = 1 
                    WHERE Id = @NotificationId 
                      AND (TargetUserId IS NULL OR LOWER(LTRIM(RTRIM(TargetUserId))) = LOWER(LTRIM(RTRIM(@UserId))));";

                db.Execute(sql, new { NotificationId = notificationId, UserId = userId?.Trim() });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error marking notification '{NotificationId}' as read for user '{UserId}'", notificationId, userId);
            }
        }

        public void MarkAllAsRead(string userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    UPDATE AppNotifications 
                    SET IsRead = 1 
                    WHERE (TargetUserId IS NULL OR LOWER(LTRIM(RTRIM(TargetUserId))) = LOWER(LTRIM(RTRIM(@UserId))))
                      AND IsRead = 0;";

                db.Execute(sql, new { UserId = userId?.Trim() });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error marking all notifications as read for user '{UserId}'", userId);
            }
        }
    }
}

