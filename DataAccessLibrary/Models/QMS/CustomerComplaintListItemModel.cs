using System;

namespace DataAccessLibrary.Models.QMS
{
    public class CustomerComplaintListItemModel
    {
        public int EntryID { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string ComplaintNo { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public byte? Type { get; set; } // 0: Customer Complaint, 1: Feedback / Observation, 2: Internal / External Audit
        public string? NCNo { get; set; }
        public string? CAPANo { get; set; }
        public string? CreditNoteNo { get; set; }
        public bool PDFAttached { get; set; }
        public int AttachmentCount { get; set; }
        public bool IsSelected { get; set; }

        public string TypeName => Type switch
        {
            0 => "Customer Complaint",
            1 => "Feedback / Observation",
            2 => "Internal / External Audit",
            _ => "Unknown"
        };

        public string TypeBadgeClass => Type switch
        {
            0 => "bg-danger-subtle text-danger border-danger-subtle",
            1 => "bg-info-subtle text-info border-info-subtle",
            2 => "bg-warning-subtle text-warning-emphasis border-warning-subtle",
            _ => "bg-secondary-subtle text-secondary"
        };
    }

    public class CustomerComplaintFilterDto
    {
        public string? CustCode { get; set; }
        public string? Country { get; set; }
        public string? ComplaintNo { get; set; }
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public bool IncludeComplaints { get; set; } = true;
        public bool IncludeFeedback { get; set; } = true;
        public bool IncludeAudits { get; set; } = true;
    }
}

