using System;

namespace DataAccessLibrary.Models.QMS
{
    public class SupplierEvaluationModel
    {
        public int EntryID { get; set; }
        public string AccNo { get; set; } = string.Empty;
        public string? AccTitle { get; set; }
        public DateTime? EntryDT { get; set; } = DateTime.Today;

        // Capacity and Operations
        public string? EmpNo { get; set; }
        public string? MachineNo { get; set; }
        public string? ToolNo { get; set; }
        public string? ProductLine { get; set; }
        public string? CoreCompetencies { get; set; }
        public string? Capacity { get; set; }
        public string? MajorCustomers { get; set; }
        public string? Remarks1 { get; set; }

        // Technical & Facility Checklist
        public bool Generator { get; set; }
        public bool SamplingHandling { get; set; }
        public bool UPS { get; set; }
        public bool Cleaniness { get; set; }
        public bool InspRecMaintain { get; set; }
        public bool TracebilityControl { get; set; }
        public bool CalibrationSystem { get; set; }
        public bool DrawingComprehension { get; set; }
        public string? Remarks2 { get; set; }

        // Additional for Vendors
        public bool ISO9000 { get; set; }
        public bool CustComplaintHanding { get; set; }
        public bool StaffTrainingSystem { get; set; }

        // Evaluation & Approvals
        public int? MarketingReputation { get; set; } = -1; // 0 = Poor, 1 = Avg, 2 = Good, -1 = None
        public int? TechnicalCapability { get; set; } = -1; // 0 = Poor, 1 = Avg, 2 = Good, -1 = None
        public string? TimeApproval { get; set; }
        public string? Remarks { get; set; }

        // Sign-off
        public string? PreparedBy { get; set; }
        public string? EvaluatedBy { get; set; }
        public string? ApprovedBy { get; set; }
        public string? PDFFileName { get; set; }
    }

    public class SupplierEvaluationListItemDto
    {
        public int EntryID { get; set; }
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public DateTime? EntryDT { get; set; }
        public string? TimeApproval { get; set; }
        public int? MarketingReputation { get; set; }
        public int? TechnicalCapability { get; set; }
        public string? Capacity { get; set; }
        public string? ProductLine { get; set; }
        public string? Remarks1 { get; set; }
        public string? Remarks { get; set; }
        public int AttachmentCount { get; set; }

        public string MarketingReputationText => MarketingReputation switch
        {
            0 => "Poor",
            1 => "Avg",
            2 => "Good",
            _ => "-"
        };

        public string MarketingBadgeClass => MarketingReputation switch
        {
            0 => "bg-danger-subtle text-danger border border-danger-subtle",
            1 => "bg-warning-subtle text-warning-emphasis border border-warning-subtle",
            2 => "bg-success-subtle text-success border border-success-subtle",
            _ => "bg-light text-secondary border"
        };

        public string TechnicalCapabilityText => TechnicalCapability switch
        {
            0 => "Poor",
            1 => "Avg",
            2 => "Good",
            _ => "-"
        };

        public string TechnicalBadgeClass => TechnicalCapability switch
        {
            0 => "bg-danger-subtle text-danger border border-danger-subtle",
            1 => "bg-warning-subtle text-warning-emphasis border border-warning-subtle",
            2 => "bg-success-subtle text-success border border-success-subtle",
            _ => "bg-light text-secondary border"
        };
    }

    public class SupplierEvaluationFilterDto
    {
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public string? AccNo { get; set; }
        public string? SearchText { get; set; }
    }

    public class SupplierDropdownDto
    {
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public int Vender { get; set; } // 1 = Vender, 0 = Maker

        public string DisplayText => $"[{AccNo}] {AccTitle} " + (Vender == 1 ? "(Vendor)" : "(Maker)");
    }
}

