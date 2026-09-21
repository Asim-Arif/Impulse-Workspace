using System;

namespace Impulse.Services.IntraOffice
{
    public class StickyNoteNotificationService
    {
        // Event passes: UserId, NoteId, Content, Color
        public event Action<string, int, string, string>? OnReminderReceived;

        public void NotifyReminder(string userId, int noteId, string content, string color)
        {
            OnReminderReceived?.Invoke(userId, noteId, content, color);
        }
    }
}
