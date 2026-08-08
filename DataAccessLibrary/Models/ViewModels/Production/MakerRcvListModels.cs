using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class MakerRcvListItem
    {
        // Primary keys
        public long EntryID { get; set; }          // VendReceived.EntryID
        public long VRD_EntryID { get; set; }      // VendRcvdDetail.EntryID

        // Maker
        public int VendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public string DisplayMaker => !string.IsNullOrEmpty(VendID1)
            ? $"{{ {VendID1} }} {VenderName}"
            : VenderName;

        // Header
        public string RecieptID { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;  // Process description

        // Article
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string TipSize { get; set; } = string.Empty;
        public string ItemSize { get; set; } = string.Empty;
        public string SizeUnit { get; set; } = string.Empty;
        public string FullArticle => $"{ItemCode} {ItemName} {TipSize} {ItemSize} {SizeUnit}".Trim();

        // Quantities
        public decimal RcvdQty { get; set; }
        public decimal Wastage { get; set; }
        public decimal LostQty { get; set; }
        public decimal WastageTotal => Wastage + LostQty;
        public decimal ReWorkQty { get; set; }
        public decimal IssQty { get; set; }

        // Status flags
        public bool ReqAuth { get; set; }
        public long? VRD_RefID { get; set; }   // Non-null means Shipped
        public string Status
        {
            get
            {
                if (VRD_RefID.HasValue && VRD_RefID.Value != 0) return "Shipped";
                return ReqAuth ? "Under Auth." : "Authorized";
            }
        }

        // Lot / Order
        public string LotNo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string MasterPONo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public long? IssuanceRefID { get; set; }   // VendIssued.EntryID linked

        // Classification
        public string CatID { get; set; } = string.Empty;
        public int GroupID { get; set; }
        public string CustCode { get; set; } = string.Empty;

        // Row color flags
        public bool ReWorkLot { get; set; }
        public string? ComplaintItemID { get; set; }
        public string? CompItemCode { get; set; }  // null = red row

        // Billing
        public long? MPB_D_EntryID { get; set; }
        public string BillNo { get; set; } = string.Empty;
        public bool IsBilled => MPB_D_EntryID.HasValue;

        // Extra
        public string Employees { get; set; } = string.Empty;
        public DateTime? EntryDT { get; set; }
    }

    public class MakerRcvListFilter
    {
        public int DateRangeIndex { get; set; } = 2; // 0=Today, 1=Last 15d, 2=Last 30d, 3=Last 60d, 4=Last 90d, 5=Custom
        public DateTime DtFrom { get; set; } = DateTime.Today.AddDays(-30);
        public DateTime DtTo { get; set; } = DateTime.Today;

        // Multi-select filters
        public HashSet<int> MakerIds { get; set; } = new HashSet<int>();
        public HashSet<string> ItemCatIds { get; set; } = new HashSet<string>();
        public HashSet<int> ItemGroupIds { get; set; } = new HashSet<int>();
        public HashSet<int> ProcessIds { get; set; } = new HashSet<int>();
        public HashSet<string> CustomerCodes { get; set; } = new HashSet<string>();

        // Single-select filters
        public string ItemId { get; set; } = "0";

        // Text filters (bypass date range)
        public string LotNo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string MasterPONo { get; set; } = string.Empty;
        public string InActiveDays { get; set; } = string.Empty;

        // Checkbox filters
        public bool ShowMasterPOOnly { get; set; }
        public bool RepairLots { get; set; }
        public bool RejectionLots { get; set; }
        public bool RegularLotsOnly { get; set; }
        public bool RcvdButNotIssued { get; set; }
        public bool BilledOnly { get; set; }
        public bool NonBilledOnly { get; set; }
        public bool ShowLastOnly { get; set; }
    }

    public class ChangeOrderNoRequest
    {
        public long EntryID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string FromOrderNo { get; set; } = string.Empty;
        public string ToOrderNo { get; set; } = string.Empty;
        public string ToItemCode { get; set; } = string.Empty;
        public decimal OriginalQty { get; set; }
        public bool TransferToStockOrder { get; set; }
    }

    public class SplitLotRequest
    {
        public long EntryID { get; set; }
        public long VR_EntryID { get; set; }
        public string OriginalLotNo { get; set; } = string.Empty;
        public string FromOrderNo { get; set; } = string.Empty;
        public string SplitOrderNo { get; set; } = string.Empty;
        public string SplitItemCode { get; set; } = string.Empty;
        public decimal OriginalQty { get; set; }
        public decimal SplitQty { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }

    public class TransferSFStockRequest
    {
        public long EntryID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string FromOrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public decimal OriginalQty { get; set; }
        public decimal TransferQty { get; set; }
        public int ShelfRefID { get; set; }
        public string LocationText { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }

    public class OrderLookupItem
    {
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string DisplayName => !string.IsNullOrEmpty(InternalRefNo)
            ? $"{OrderNo} {{{InternalRefNo}}}"
            : OrderNo;
    }

    public class StoreLookupItem
    {
        public int EntryID { get; set; }
        public string StoreName { get; set; } = string.Empty;
    }

    public class ShelfLookupItem
    {
        public int EntryID { get; set; }
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public string DisplayName => $"{RackNo} {{{ShelfNo}}}";
    }

    public class ProcessOptionItem
    {
        public int ProcessID { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string DisplayName => !string.IsNullOrEmpty(Code)
            ? $"{Code} - {Description}"
            : Description;
    }

    public class InspectionParameterItem
    {
        public int EntryID { get; set; }            // PIP_RefID
        public string ParameterName { get; set; } = string.Empty;
        public string AQL { get; set; } = string.Empty;
        public string SampleSize { get; set; } = string.Empty;
        public string RejectOn { get; set; } = string.Empty;
        public string TestSpecificationNo { get; set; } = string.Empty;
        public string ActualRejection { get; set; } = "0";
        public string Status { get; set; } = "OK";
    }

    public class InspectionDataDto
    {
        public long EntryID { get; set; }          // VRDI EntryID (if already saved)
        public long VRD_RefID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string ItemSize { get; set; } = string.Empty;
        public string SizeUnit { get; set; } = string.Empty;
        public string TipSize { get; set; } = string.Empty;
        public decimal RcvdQty { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ProcessDescription { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public DateTime? ReceivingDT { get; set; }

        public DateTime InspectionDT { get; set; } = DateTime.Today;
        public bool LotStatus { get; set; } = false; // false = Pass, true = Fail
        public string Disposation { get; set; } = string.Empty;
        public string Comments { get; set; } = string.Empty;

        public List<ProcessOptionItem> Processes { get; set; } = new List<ProcessOptionItem>();
        public List<InspectionParameterItem> Parameters { get; set; } = new List<InspectionParameterItem>();
        public List<string> TemperValues { get; set; } = new List<string>();
    }

    public class SaveInspectionRequest
    {
        public long VRD_RefID { get; set; }
        public DateTime DT { get; set; } = DateTime.Today;
        public bool LotStatus { get; set; } // false = Pass, true = Fail
        public string Disposation { get; set; } = string.Empty;
        public string Comments { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;

        public List<InspectionParameterItem> Parameters { get; set; } = new List<InspectionParameterItem>();
        public List<string> TemperValues { get; set; } = new List<string>();
    }
}


