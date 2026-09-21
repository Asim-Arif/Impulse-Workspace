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
        public bool EmailMessageSent { get; set; } = false;
        public bool IsRead { get; set; } = false;
        public string? AdditionalAssigneeIds { get; set; }
        public string? AssignedToNames { get; set; }
        public DateTime? StartedAt { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? UpdatedAt { get; set; }
        public DateTime? CompletedAt { get; set; }

        public IntraUserProfile? Assignee => !string.IsNullOrEmpty(AssignedTo) ? new IntraUserProfile { UserName = AssignedTo, FullUserName = AssigneeName, Designation = AssigneeDesignation } : null;
        public IntraUserProfile? Assigner => !string.IsNullOrEmpty(AssignedBy) ? new IntraUserProfile { UserName = AssignedBy, FullUserName = AssignerName } : null;

        public List<TaskComment> Comments { get; set; } = new List<TaskComment>();
        public List<TaskAttachment> Attachments { get; set; } = new List<TaskAttachment>();
        public List<TaskStatusHistory> StatusHistories { get; set; } = new List<TaskStatusHistory>();
    }

    public class TaskStatusHistory
    {
        public int Id { get; set; }
        public int TaskId { get; set; }
        public TaskItemStatus FromStatus { get; set; }
        public TaskItemStatus ToStatus { get; set; }
        public TaskItemStatus OldStatus { get => FromStatus; set => FromStatus = value; }
        public TaskItemStatus NewStatus { get => ToStatus; set => ToStatus = value; }
        public string? ChangedById { get; set; }
        public ApplicationUser? ChangedBy { get; set; }
        public long DurationSeconds { get; set; }
        public DateTime ChangedAt { get; set; } = DateTime.UtcNow;
        public string? Notes { get; set; }
    }

    public class TaskComment
    {
        public int Id { get; set; }
        public int TaskId { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string? UserName { get; set; }
        public string? UserFullName { get; set; }
        public IntraUserProfile? User => !string.IsNullOrEmpty(UserId) ? new IntraUserProfile { UserName = UserId, FullUserName = UserFullName } : null;
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
