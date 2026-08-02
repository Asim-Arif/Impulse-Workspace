using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class PendingAuthorizeReceivedItemModel
    {
        public long EntryID { get; set; } // VendRcvdDetail.EntryID
        public long RefID { get; set; }   // VendReceived.EntryID
        public string RecieptID { get; set; } = string.Empty;
        public long VendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public string Supervisor { get; set; } = string.Empty;
        public string DisplayVendor => !string.IsNullOrEmpty(VenderName) ? VenderName : Supervisor;
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemDescription { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string Process { get; set; } = string.Empty;
        public decimal RcvdQty { get; set; }
        public decimal Wastage { get; set; }
        public decimal LostQty { get; set; }
        public decimal ReWorkQty { get; set; }
        public decimal RepairAmt { get; set; }
        public decimal Rate { get; set; }
        public long Issue_RefID { get; set; }
        public long Issuance_RefID { get; set; }
        public double RcvdWeight { get; set; }
        public string Unit { get; set; } = "Pcs";
        public bool IsChecked { get; set; } = true;
        public string Insp_EmpID { get; set; } = string.Empty;

        // In-memory breakdown collections
        public List<WastageBreakdownItemModel> WastageBreakdowns { get; set; } = new List<WastageBreakdownItemModel>();
        public List<ReWorkBreakdownItemModel> ReWorkBreakdowns { get; set; } = new List<ReWorkBreakdownItemModel>();
    }

    public class WastageTypeLookupModel
    {
        public long EntryID { get; set; }
        public string Description { get; set; } = string.Empty;
        public int WastageType { get; set; } // 1=Vendor, 2=Originator, 3=Selection
    }

    public class RepairTypeLookupModel
    {
        public int EntryID { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public class WastageBreakdownItemModel
    {
        public long EntryID { get; set; }
        public long VRD_RefID { get; set; }
        public long Wastage_RefID { get; set; }
        public string WastageTypeName { get; set; } = string.Empty;
        public decimal Qty { get; set; }
        public int WastageType { get; set; } // 1=Vendor, 2=Originator, 3=Selection
        public long ReturnTo_VRD_RefID { get; set; }
        public long VendID { get; set; }
        public string VendName { get; set; } = string.Empty;
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
    }

    public class ReWorkBreakdownItemModel
    {
        public long EntryID { get; set; }
        public long VRD_RefID { get; set; }
        public int Repair_RefID { get; set; }
        public string RepairTypeName { get; set; } = string.Empty;
        public decimal Qty { get; set; }
        public int ReturnProcessID { get; set; }
        public string ReturnProcessName { get; set; } = string.Empty;
    }

    public class AuthorizeReceivedFilterModel
    {
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public long VendID { get; set; } = 0;
        public int ProcessID { get; set; } = 0;
        public string SearchText { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
    }
}
