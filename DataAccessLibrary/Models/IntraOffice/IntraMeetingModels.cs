using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class Meeting
    {
        public int Id { get; set; }
        public string Title { get; set; } = string.Empty;
        public string OrganizerId { get; set; } = string.Empty;
        public string? OrganizerName { get; set; }
        public string? MeetingUrl { get; set; }
        public DateTime ScheduledStartTime { get; set; }
        public DateTime? ScheduledEndTime { get; set; }
        public MeetingStatus Status { get; set; } = MeetingStatus.Scheduled;
        public string? MeetingMinutes { get; set; }
        public bool IsReminderSent { get; set; } = false;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public List<MeetingParticipant> Participants { get; set; } = new List<MeetingParticipant>();
    }

    public class MeetingParticipant
    {
        public int MeetingId { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string? UserName { get; set; }
        public string? FullName { get; set; }
        public string? Designation { get; set; }
        public bool HasAttended { get; set; } = false;
    }
}
