using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class Announcement
    {
        public int Id { get; set; }
        public string Title { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public AnnouncementPriority Priority { get; set; } = AnnouncementPriority.Normal;
        public string CreatedBy { get; set; } = string.Empty;
        public string? CreatorName { get; set; }
        public string? CreatorDesignation { get; set; }
        public string? DepartmentId { get; set; }
        public string? DepartmentName { get; set; }
        public bool IsPinned { get; set; } = false;
        public bool IsActive { get; set; } = true;
        public bool IsAcknowledged { get; set; } = false;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? UpdatedAt { get; set; }
        public DateTime? ExpiresAt { get; set; }

        public IntraUserProfile? Author => !string.IsNullOrEmpty(CreatedBy) ? new IntraUserProfile { UserName = CreatedBy, FullUserName = CreatorName ?? CreatedBy, Designation = CreatorDesignation } : null;
        public IntraDepartmentInfo? Department => !string.IsNullOrEmpty(DepartmentName) ? new IntraDepartmentInfo { Name = DepartmentName } : null;

        public int AcknowledgmentCount { get; set; }
        public bool IsAcknowledgedByCurrentUser { get; set; }
        public List<AnnouncementAttachment> Attachments { get; set; } = new List<AnnouncementAttachment>();
        public List<AnnouncementAcknowledgment> Acknowledgments { get; set; } = new List<AnnouncementAcknowledgment>();
    }

    public class AnnouncementAttachment
    {
        public int Id { get; set; }
        public int AnnouncementId { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string? ContentType { get; set; }
        public DateTime UploadedAt { get; set; } = DateTime.UtcNow;
    }

    public class AnnouncementAcknowledgment
    {
        public int Id { get; set; }
        public int AnnouncementId { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string? UserName { get; set; }
        public string? FullName { get; set; }
        public IntraUserProfile? User => !string.IsNullOrEmpty(UserId) ? new IntraUserProfile { UserName = UserId, FullUserName = FullName ?? UserName ?? UserId } : null;
        public DateTime AcknowledgedAt { get; set; } = DateTime.UtcNow;
    }
}
