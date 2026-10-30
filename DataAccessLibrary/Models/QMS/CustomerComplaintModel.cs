using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.QMS
{
    public class CustomerComplaintModel
    {
        public int EntryID { get; set; }
        public int? Company_RefID { get; set; } = 1;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string ComplaintNo { get; set; } = string.Empty;
        public DateTime? DT { get; set; } = DateTime.Today;
        public string? Unit { get; set; }
        public string? ComplaintBy { get; set; }
        public string? Designation { get; set; }
        public byte? Penalty { get; set; } = 0; // 0: None, 1: Free Placement, 2: Credit Note, 3: Investigation
        public int? PenaltyQty { get; set; }
        public float? PenaltyAmount { get; set; }
        public string? PenaltyCurrency { get; set; }
        public string? EmailTo { get; set; }
        public string? EmailCc { get; set; }
        public byte? Type { get; set; } = 0; // 0: Customer Complaint, 1: Feedback, 2: Audit
        public bool? InternalAudit { get; set; } = true;
        public DateTime? DTNextAudit { get; set; }
        public string? Findings { get; set; }
        public string? Description { get; set; }
        public string? Reason { get; set; }
        public string? ActionRequested { get; set; }
        public string? NCNo { get; set; }
        public string? CAPANo { get; set; }
        public string? CreditNoteNo { get; set; }
        public string? UserName { get; set; }
        public string? MachineName { get; set; }
        public DateTime? DTEntry { get; set; }

        public List<CustomerComplaintDetailModel> Details { get; set; } = new();
    }

    public class CustomerComplaintDetailModel
    {
        public int EntryID { get; set; }
        public int RefID { get; set; }
        public string? InvoiceNo { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string? ItemCode { get; set; }
        public int? Qty { get; set; }
        public string? EmailReference { get; set; } // Lot / Reference
        public string? ItemReason_Main { get; set; } // QMS Reason
        public string? ItemReason { get; set; } // Detailed Description
        public float? Price { get; set; }
        public bool IsDeleted { get; set; }
    }

    public class CompanyLookupDto
    {
        public int EntryID { get; set; }
        public string CompanyName { get; set; } = string.Empty;
    }

    public class DepartmentLookupDto
    {
        public string DeptID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }

    public class QMSReasonLookupDto
    {
        public int EntryID { get; set; }
        public string Reason_QMS { get; set; } = string.Empty;
    }

    public class EmployeeLookupDto
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string DisplayName => $"[{EmpID}] {Name}";
    }

    public class InvoiceItemLookupDto
    {
        public string ItemID { get; set; } = string.Empty;
        public string? CompItemID { get; set; }
        public string? ItemCode { get; set; }
        public string DisplayName => !string.IsNullOrEmpty(CompItemID) ? $"{{{CompItemID}}} {ItemID}" : ItemID;
    }
}

