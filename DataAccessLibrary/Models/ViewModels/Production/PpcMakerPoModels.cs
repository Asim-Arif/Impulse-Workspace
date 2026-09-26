using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class PpcMakerPoOrderHeaderDto
    {
        public string OrderNo { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public DateTime OrderDate { get; set; }
        public DateTime? DeliveryDate { get; set; }
        public string? CustomerPoNo { get; set; }
        public int TotalPlannedPurchaseQty { get; set; }
        public int TotalPurchaseLinesCount { get; set; }
        public int PendingPurchaseLinesCount { get; set; }
        public bool HasPendingPurchases => PendingPurchaseLinesCount > 0;
    }

    public class PpcMakerPoItemRowDto
    {
        public long EntryID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string CompItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public int VendID { get; set; }
        public string MakerName { get; set; } = string.Empty;
        
        // PO Qty locked strictly matching PPC
        public int PurchaseQty { get; set; }
        
        // Editable fields
        public decimal PurchaseRate { get; set; }
        public DateTime ReturnDT { get; set; } = DateTime.Today.AddDays(30);
        public string BatchNo { get; set; } = string.Empty;
        public int? SteelType_RefID { get; set; }
        public string? SpecialInstructions { get; set; }
        public bool SampleProvided { get; set; }
        public bool DrawingProvided { get; set; }
        public bool ForgingProvided { get; set; }
        public bool SteelProvided { get; set; }

        // Fulfilled / Issued State
        public long? VendIssued_RefID { get; set; }
        public string? MasterPONo { get; set; }
        public DateTime? IssuedAt { get; set; }
        public string? IssuedBy { get; set; }

        public bool IsSelected { get; set; } = true;
        public bool IsPosted => VendIssued_RefID.HasValue && !string.IsNullOrWhiteSpace(MasterPONo);
    }

    public class GenerateMakerPoRequest
    {
        public string OrderNo { get; set; } = string.Empty;
        public string? IssEmpID { get; set; }
        public string? CountedBy { get; set; }
        public int? SteelType_RefID { get; set; }
        public List<PpcMakerPoItemRowDto> Lines { get; set; } = new();
        public string UserName { get; set; } = "Purchaser";
        public int UserId { get; set; } = 1;
        public string MachineName { get; set; } = "WEB-CLIENT";
    }

    public class GenerateMakerPoResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public List<string> GeneratedMasterPoNumbers { get; set; } = new();
        public int TotalRowsCreated { get; set; }
    }
}
