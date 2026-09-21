using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

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
        private readonly List<AppNotification> _notifications = new();
        private readonly object _lock = new();

        public async Task SendNotificationAsync(AppNotification notification)
        {
            lock (_lock)
            {
                _notifications.Insert(0, notification);
                if (_notifications.Count > 200)
                {
                    _notifications.RemoveAt(_notifications.Count - 1);
                }
            }

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
            lock (_lock)
            {
                return _notifications
                    .Where(n => string.IsNullOrEmpty(n.TargetUserId) || n.TargetUserId == userId)
                    .Take(30)
                    .ToList();
            }
        }

        public int GetUnreadCount(string userId)
        {
            lock (_lock)
            {
                return _notifications
                    .Count(n => (string.IsNullOrEmpty(n.TargetUserId) || n.TargetUserId == userId) && !n.IsRead);
            }
        }

        public void MarkAsRead(string notificationId, string userId)
        {
            lock (_lock)
            {
                var n = _notifications.FirstOrDefault(x => x.Id == notificationId);
                if (n != null)
                {
                    n.IsRead = true;
                }
            }
        }

        public void MarkAllAsRead(string userId)
        {
            lock (_lock)
            {
                foreach (var n in _notifications.Where(n => string.IsNullOrEmpty(n.TargetUserId) || n.TargetUserId == userId))
                {
                    n.IsRead = true;
                }
            }
        }
    }
}
