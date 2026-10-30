using System;

namespace DataAccessLibrary.Models.QMS
{
    public class CAPAModel
    {
        public int EntryID { get; set; }
        public string CAPANo { get; set; } = string.Empty;
        public DateTime? DT { get; set; } = DateTime.Today;

        // Step 1: Investigation & Root Cause Analysis
        public int? Reason { get; set; } = 0; // 0: Product, 1: Process, 2: System, 3: Service
        public int? System { get; set; }
        public string? Description { get; set; }
        public int? Acceptance { get; set; } = 0; // 0: Acceptance for further Action, 1: Hold for further action
        public string? TeamNominated { get; set; }
        public string? ResultOfInvestigation { get; set; }
        public DateTime? ResultDT { get; set; }

        // Step 1: Checkbox Systems
        public bool System_QM { get; set; }   // Quality Management
        public bool System_EM { get; set; }   // Environment Management
        public bool System_cGMP { get; set; } // cGMP / CFR 21 820
        public bool System_OHSM { get; set; } // Occup. Health & Safety Management
        public bool System_ITSM { get; set; } // IT Security Management
        public bool System_MDD { get; set; }  // Medical Device Directive 93/42/EEC

        // Step 2: Action Plan (Corrective & Preventive Actions)
        public int? CAReview { get; set; } // 0: Change in Corrective Action required, 1: Change in Implementation Date required, 2: Change in Follow-up Date required
        public string? CorrectiveAction { get; set; }
        public DateTime? CorrectiveActionDT { get; set; }
        public string? Remarks { get; set; }
        public string? PreventiveActions { get; set; }
        public DateTime? PreventiveActionDT { get; set; }

        // Step 3: Verification, Follow-up & Closure
        public int? ResponseEvaluation { get; set; } // 0: Verification audit is required, 1: Verification of submitted documentation or self declaration prior to closure, 2: Proposal of corrective action acceptable. Verify effectiveness.
        public string? FollowUp { get; set; }
        public DateTime? CorrectiveActionFollowupDT { get; set; }
        public DateTime? PreventiveActionFollowUpDT { get; set; }
        public DateTime? CAPACloseDT { get; set; }

        // Associated Customer Complaint Details (Read-Only)
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
        public string? ComplaintReason { get; set; }

        public string PenaltyDisplay => Penalty switch
        {
            1 => "Free Placement",
            2 => "Credit Note",
            3 => "Investigation",
            _ => "None"
        };

        // Associated Non-Compliance Details (Read-Only)
        public int? NCEntryID { get; set; }
        public string? NCNo { get; set; }
        public string? NCType { get; set; }
        public DateTime? NCDate { get; set; }
        public string? NCSource { get; set; }
        public string? NCDepartment { get; set; }
        public string? NCDescription { get; set; }
    }
}

