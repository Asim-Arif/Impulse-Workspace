using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class WorkflowTemplate
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string RequestType { get; set; } = string.Empty;
        public bool IsActive { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? UpdatedAt { get; set; }

        public List<WorkflowStep> Steps { get; set; } = new List<WorkflowStep>();
    }

    public class WorkflowStep
    {
        public int Id { get; set; }
        public int WorkflowTemplateId { get; set; }
        public int StepOrder { get; set; }
        public string StepName { get; set; } = string.Empty;
        public ApproverType ApproverType { get; set; } = ApproverType.Manager;
        public string? SpecificUserId { get; set; }
        public string? RoleName { get; set; }
        public int RequiredAction { get; set; }
        public int SLAHours { get; set; } = 24;
    }

    public class ApprovalRequest
    {
        public int Id { get; set; }
        public string RequestNumber { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string? Description { get; set; }
        public int WorkflowTemplateId { get; set; }
        public string? WorkflowTemplateName { get; set; }
        public string RequesterId { get; set; } = string.Empty;
        public string? RequesterName { get; set; }
        public string? DepartmentId { get; set; }
        public string? DepartmentName { get; set; }
        public int Priority { get; set; } = 1;
        public int Status { get; set; } = 0;
        public DateTime RequestedDate { get; set; } = DateTime.UtcNow;
        public DateTime? EffectiveDate { get; set; }
        public decimal? Amount { get; set; }
        public string? CostCenter { get; set; }
        public string? Project { get; set; }
        public int CurrentStepOrder { get; set; } = 1;

        public List<ApprovalRequestAttachment> Attachments { get; set; } = new List<ApprovalRequestAttachment>();
        public List<ApprovalRequestHistory> Histories { get; set; } = new List<ApprovalRequestHistory>();
    }

    public class ApprovalRequestAttachment
    {
        public int Id { get; set; }
        public int ApprovalRequestId { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string? ContentType { get; set; }
        public DateTime UploadedAt { get; set; } = DateTime.UtcNow;
    }

    public class ApprovalRequestHistory
    {
        public int Id { get; set; }
        public int ApprovalRequestId { get; set; }
        public int StepOrder { get; set; }
        public int ActionTaken { get; set; }
        public string ActionById { get; set; } = string.Empty;
        public string? ActionByName { get; set; }
        public DateTime ActionDate { get; set; } = DateTime.UtcNow;
        public string? Comments { get; set; }
    }

    public class ApprovalDelegation
    {
        public int Id { get; set; }
        public string DelegatorId { get; set; } = string.Empty;
        public string DelegateeId { get; set; } = string.Empty;
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public bool IsActive { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
