using System;

namespace DataAccessLibrary.Models.Setup
{
    public class UserPushSubscriptionModel
    {
        public int Id { get; set; }
        public int? UserID { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string Endpoint { get; set; } = string.Empty;
        public string P256dh { get; set; } = string.Empty;
        public string Auth { get; set; } = string.Empty;
        public string? DeviceName { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? LastUsedAt { get; set; }
    }

    public class PushSubscriptionDto
    {
        public string Endpoint { get; set; } = string.Empty;
        public PushSubscriptionKeys Keys { get; set; } = new();
        public string? DeviceName { get; set; }
        public string? UserId { get; set; }
        public string? UserName { get; set; }
    }

    public class PushSubscriptionKeys
    {
        public string P256dh { get; set; } = string.Empty;
        public string Auth { get; set; } = string.Empty;
    }
}
