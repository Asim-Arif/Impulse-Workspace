using System;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class UserPresence
    {
        public string UserId { get; set; } = string.Empty;
        public UserStatus Status { get; set; } = UserStatus.Offline;
        public DateTime LastSeen { get; set; } = DateTime.UtcNow;
        public string? ConnectionId { get; set; }
    }

    public class StickyNote
    {
        public int Id { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public string Color { get; set; } = "#fffa65";
        public int XPos { get; set; } = 100;
        public int YPos { get; set; } = 100;
        public DateTime? ReminderTime { get; set; }
        public bool IsReminderSent { get; set; } = false;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    public class GameScore
    {
        public int Id { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string GameName { get; set; } = string.Empty;
        public int Score { get; set; }
        public DateTime PlayedAt { get; set; } = DateTime.UtcNow;
    }
}
