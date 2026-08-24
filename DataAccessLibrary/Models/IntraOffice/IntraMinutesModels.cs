using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class MinuteType
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public bool IsActive { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    public class MinuteApproval
    {
        public int Id { get; set; }
        public DateTime Date { get; set; } = DateTime.Today;
        public string No { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Points { get; set; } = string.Empty;
        public string? ForwardToUserId { get; set; }
        public string? ForwardToUserName { get; set; }
        public string? Currency { get; set; } = "PKR";
        public decimal? TotalAmount { get; set; }
        public decimal? AdvancePercentage { get; set; }
        public decimal? AdvanceAmount { get; set; }
        public bool IsUrgent { get; set; } = false;
        public bool CloseByInitiator { get; set; } = false;
        public string CreatedByUserId { get; set; } = string.Empty;
        public string? CreatedByUserName { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public string Status { get; set; } = "Pending";
        public string? SignaturePath { get; set; }
        public decimal? PurchaseOrderValue { get; set; }
        public decimal? PurchaseAdvanceRecommend { get; set; }
        public decimal? PurchaseApprovedAmount { get; set; }
        public decimal? RequestedStockQty { get; set; }
        public decimal? CurrentStockQty { get; set; }
        public decimal? ApprovedStockQty { get; set; }
        public string? HRLeaveType { get; set; }
        public string? FinancialType { get; set; }

        public List<MinuteAttachment> Attachments { get; set; } = new List<MinuteAttachment>();
        public List<MinuteWorkflowHistory> WorkflowHistory { get; set; } = new List<MinuteWorkflowHistory>();
    }

    public class MinuteAttachment
    {
        public int Id { get; set; }
        public int MinuteApprovalId { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string? ContentType { get; set; }
        public DateTime UploadedAt { get; set; } = DateTime.UtcNow;
    }

    public class MinuteWorkflowHistory
    {
        public int Id { get; set; }
        public int MinuteApprovalId { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string? UserName { get; set; }
        public string? FullName { get; set; }
        public string? Designation { get; set; }
        public string ActionTaken { get; set; } = string.Empty;
        public string? Remarks { get; set; }
        public string? SignaturePath { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    public class MinuteFilter
    {
        public DateTime? FromDate { get; set; }
        public DateTime? ToDate { get; set; }
        public string? No { get; set; }
        public string? Type { get; set; }
        public string? Status { get; set; }
        public string? UserId { get; set; }
        public string? SearchTerm { get; set; }
    }
}
