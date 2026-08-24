using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class TaskItem
    {
        public int Id { get; set; }
        public string Title { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string? AssignedTo { get; set; }
        public string? AssigneeName { get; set; }
        public string? AssigneeDesignation { get; set; }
        public string AssignedBy { get; set; } = string.Empty;
        public string? AssignerName { get; set; }
        public string? DepartmentId { get; set; }
        public string? DepartmentName { get; set; }
        public TaskPriority Priority { get; set; } = TaskPriority.Normal;
        public TaskItemStatus Status { get; set; } = TaskItemStatus.Pending;
        public DateTime? DueDate { get; set; }
        public bool WhatsAppMessageSent { get; set; } = false;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? UpdatedAt { get; set; }
        public DateTime? CompletedAt { get; set; }

        public List<TaskComment> Comments { get; set; } = new List<TaskComment>();
        public List<TaskAttachment> Attachments { get; set; } = new List<TaskAttachment>();
    }

    public class TaskComment
    {
        public int Id { get; set; }
        public int TaskId { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string? UserName { get; set; }
        public string? UserFullName { get; set; }
        public string Content { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    public class TaskAttachment
    {
        public int Id { get; set; }
        public int TaskId { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string? ContentType { get; set; }
        public DateTime UploadedAt { get; set; } = DateTime.UtcNow;
    }
}
