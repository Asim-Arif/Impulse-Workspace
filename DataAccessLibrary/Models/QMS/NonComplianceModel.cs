using System;

namespace DataAccessLibrary.Models.QMS
{
    public class NonComplianceModel
    {
        public int EntryID { get; set; }
        public string NCNo { get; set; } = string.Empty;
        public string? NCType { get; set; } = "Minor";
        public DateTime? DT { get; set; } = DateTime.Today;
        public string? Source { get; set; } = "Customer Complaint";
        public string? Department { get; set; }
        public string? Description { get; set; }

        // Associated Source / Customer Complaint Details (Read-Only)
        public int? ComplaintEntryID { get; set; }
        public string? CustCode { get; set; }
        public string? Country { get; set; }
        public string? ComplaintNo { get; set; }
        public DateTime? ComplaintDate { get; set; }
        public string? ComplaintBy { get; set; }
        public string? Designation { get; set; }
        public byte? Penalty { get; set; }
        public int? PenaltyQty { get; set; }
        public float? PenaltyAmount { get; set; }
        public string? PenaltyCurrency { get; set; }

        public string PenaltyDisplay => Penalty switch
        {
            1 => "Free Placement",
            2 => "Credit Note",
            3 => "Investigation",
            _ => "None"
        };
    }
}

