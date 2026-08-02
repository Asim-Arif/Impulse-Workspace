using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class MasterPOSearchResultModel
    {
        public long EntryID { get; set; }
        public string RecieptID { get; set; } = string.Empty;
        public long VendID { get; set; }
        public string VendorName { get; set; } = string.Empty;
        public string VendID1 { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public bool Authorized { get; set; }
        public bool Closed { get; set; }
        public string IssEmpID { get; set; } = string.Empty;
        public bool IsFactoryMaker { get; set; }
    }

    public class MasterPOLineItemModel
    {
        public long VendIssdDetailEntryID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public byte[]? ItemPic { get; set; }
        public string? ItemPicBase64 => ItemPic != null && ItemPic.Length > 0 ? $"data:image/jpeg;base64,{Convert.ToBase64String(ItemPic)}" : null;
        public decimal IssQty { get; set; }
        public decimal PrevRcvdQty { get; set; }
        public decimal PendingQty => IssQty - PrevRcvdQty;
        public decimal Rate { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public decimal ReceivingQty { get; set; }
        public string CountedBy { get; set; } = string.Empty;
        public int ReWorkLot { get; set; }
        public int RepairType { get; set; }
        public int ReturnProcessID { get; set; }
        public int Priority { get; set; }
    }

    public class CreatePOReceivingHeaderModel
    {
        public long IssuanceRefID { get; set; }
        public long VendID { get; set; }
        public int ProcessID { get; set; }
        public DateTime DT { get; set; } = DateTime.Today;
        public bool OverTime { get; set; } = false;
        public string TemperValue { get; set; } = string.Empty;
        public string MillCertNo { get; set; } = string.Empty;
        public List<string> FactoryEmpIDs { get; set; } = new List<string>();
    }

    public class CreatePOReceivingLineModel
    {
        public long VendIssdDetailEntryID { get; set; }
        public string ItemCode { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public decimal Rate { get; set; }
        public decimal RcvdQty { get; set; }
        public string CountedBy { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public int ReWorkLot { get; set; } = 0;
        public int RepairType { get; set; } = 0;
        public int ProcessID { get; set; }
    }
}
