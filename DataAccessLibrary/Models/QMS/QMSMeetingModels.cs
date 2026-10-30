using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.QMS
{
    public class QMSMeetingListItemModel
    {
        public int EntryID { get; set; }
        public string? Code { get; set; }
        public string? Type { get; set; }
        public DateTime? DT { get; set; }
        public string? Place { get; set; }
        public string? Duration { get; set; }
        public string? Agenda { get; set; }
        public string? Discussion { get; set; }
        public string? Keywords { get; set; }
        public int AttendeeCount { get; set; }
        public int AttachmentCount { get; set; }
        public string? UserName { get; set; }
        public DateTime? DTEntry { get; set; }
    }

    public class QMSMeetingModel
    {
        public int EntryID { get; set; }
        public string? Code { get; set; }
        public string? Type { get; set; }
        public DateTime? DT { get; set; } = DateTime.Now;
        public string? Place { get; set; }
        public string? Duration { get; set; }
        public string? Agenda { get; set; }
        public string? Discussion { get; set; }
        public string? Keywords { get; set; }
        public string? UserName { get; set; }
        public string? MachineName { get; set; }
        public DateTime? DTEntry { get; set; }
        public List<string> Attendees { get; set; } = new();
    }

    public class QMSMeetingFilterDto
    {
        public string? SearchText { get; set; }
        public string? MeetingType { get; set; }
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
    }
}

