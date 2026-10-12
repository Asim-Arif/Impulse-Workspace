using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
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
        Lead,
        Stock
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
        List<AppNotification> GetRecentNotifications(string userId, IEnumerable<string>? aliases = null);
        int GetUnreadCount(string userId, IEnumerable<string>? aliases = null);
        void MarkAsRead(string notificationId, string userId, IEnumerable<string>? aliases = null);
        void MarkAllAsRead(string userId, IEnumerable<string>? aliases = null);
    }

    public class AppNotificationService : IAppNotificationService
    {
        public event Func<AppNotification, Task>? OnNotificationReceived;
        private readonly string _connectionString;
        private readonly ILogger<AppNotificationService> _logger;
        private readonly IServiceScopeFactory _scopeFactory;

        public AppNotificationService(
            IConfiguration configuration,
            ILogger<AppNotificationService> logger,
            IServiceScopeFactory scopeFactory)
        {
            _logger = logger;
            _scopeFactory = scopeFactory;
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        private static List<string> BuildIdentifierList(string userId, IEnumerable<string>? aliases)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(userId))
            {
                set.Add(userId.Trim().ToLowerInvariant());
            }
            if (aliases != null)
            {
                foreach (var a in aliases)
                {
                    if (!string.IsNullOrWhiteSpace(a))
                    {
                        set.Add(a.Trim().ToLowerInvariant());
                    }
                }
            }
            return set.ToList();
        }

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

            // Real-time parallel dispatch to active client circuits (with per-subscriber timeout)
            if (OnNotificationReceived != null)
            {
                var delegates = OnNotificationReceived.GetInvocationList();
                _ = Task.Run(async () =>
                {
                    var tasks = delegates.Select(async d =>
                    {
                        try
                        {
                            var handler = (Func<AppNotification, Task>)d;
                            var handlerTask = handler(notification);
                            var timeoutTask = Task.Delay(4000);
                            if (await Task.WhenAny(handlerTask, timeoutTask) == handlerTask)
                            {
                                await handlerTask;
                            }
                            else
                            {
                                _logger.LogWarning("Notification dispatch handler timed out (4s) for circuit subscriber on notification #{Id}", notification.Id);
                            }
                        }
                        catch
                        {
                            // Ignore disconnected circuit errors
                        }
                    });
                    await Task.WhenAll(tasks);
                });
            }

            // Web Push notification to mobile/background devices
            if (!notification.IsReadReceipt)
            {
                _ = Task.Run(async () =>
                {
                    try
                    {
                        using var scope = _scopeFactory.CreateScope();
                        var webPushService = scope.ServiceProvider.GetService<Impulse.Services.Notifications.IWebPushNotificationService>();
                        if (webPushService != null)
                        {
                            await webPushService.SendNotificationAsync(notification);
                        }
                    }
                    catch (Exception pushEx)
                    {
                        _logger.LogWarning(pushEx, "Failed to send web push for notification #{Id}", notification.Id);
                    }
                });
            }
        }

        public List<AppNotification> GetRecentNotifications(string userId, IEnumerable<string>? aliases = null)
        {
            try
            {
                var identifiers = BuildIdentifierList(userId, aliases);
                using var db = CreateConnection();
                string sql;
                object param;

                if (identifiers.Count > 0)
                {
                    sql = @"
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
                        WHERE (TargetUserId IS NULL OR LOWER(LTRIM(RTRIM(TargetUserId))) IN @UserIdentifiers)
                        ORDER BY CreatedAt DESC;";
                    param = new { UserIdentifiers = identifiers };
                }
                else
                {
                    sql = @"
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
                        WHERE TargetUserId IS NULL
                        ORDER BY CreatedAt DESC;";
                    param = new { };
                }

                return db.Query<AppNotification>(sql, param).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching recent notifications for user '{UserId}' from database", userId);
                return new List<AppNotification>();
            }
        }

        public int GetUnreadCount(string userId, IEnumerable<string>? aliases = null)
        {
            try
            {
                var identifiers = BuildIdentifierList(userId, aliases);
                using var db = CreateConnection();
                string sql;
                object param;

                if (identifiers.Count > 0)
                {
                    sql = @"
                        SELECT COUNT(*) 
                        FROM AppNotifications
                        WHERE (TargetUserId IS NULL OR LOWER(LTRIM(RTRIM(TargetUserId))) IN @UserIdentifiers)
                          AND IsRead = 0;";
                    param = new { UserIdentifiers = identifiers };
                }
                else
                {
                    sql = @"
                        SELECT COUNT(*) 
                        FROM AppNotifications
                        WHERE TargetUserId IS NULL AND IsRead = 0;";
                    param = new { };
                }

                return db.ExecuteScalar<int>(sql, param);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error counting unread notifications for user '{UserId}' from database", userId);
                return 0;
            }
        }

        public void MarkAsRead(string notificationId, string userId, IEnumerable<string>? aliases = null)
        {
            try
            {
                var identifiers = BuildIdentifierList(userId, aliases);
                using var db = CreateConnection();
                string sql;
                object param;

                if (identifiers.Count > 0)
                {
                    sql = @"
                        UPDATE AppNotifications 
                        SET IsRead = 1 
                        WHERE Id = @NotificationId 
                          AND (TargetUserId IS NULL OR LOWER(LTRIM(RTRIM(TargetUserId))) IN @UserIdentifiers);";
                    param = new { NotificationId = notificationId, UserIdentifiers = identifiers };
                }
                else
                {
                    sql = @"
                        UPDATE AppNotifications 
                        SET IsRead = 1 
                        WHERE Id = @NotificationId;";
                    param = new { NotificationId = notificationId };
                }

                db.Execute(sql, param);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error marking notification '{NotificationId}' as read for user '{UserId}'", notificationId, userId);
            }
        }

        public void MarkAllAsRead(string userId, IEnumerable<string>? aliases = null)
        {
            try
            {
                var identifiers = BuildIdentifierList(userId, aliases);
                using var db = CreateConnection();
                string sql;
                object param;

                if (identifiers.Count > 0)
                {
                    sql = @"
                        UPDATE AppNotifications 
                        SET IsRead = 1 
                        WHERE (TargetUserId IS NULL OR LOWER(LTRIM(RTRIM(TargetUserId))) IN @UserIdentifiers)
                          AND IsRead = 0;";
                    param = new { UserIdentifiers = identifiers };
                }
                else
                {
                    sql = @"
                        UPDATE AppNotifications 
                        SET IsRead = 1 
                        WHERE TargetUserId IS NULL AND IsRead = 0;";
                    param = new { };
                }

                db.Execute(sql, param);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error marking all notifications as read for user '{UserId}'", userId);
            }
        }
    }
}

