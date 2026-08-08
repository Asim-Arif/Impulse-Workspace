using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class ReWorkAvailableItem
    {
        public long EntryID { get; set; }          // VendRcvdDetailReWorkDetail.EntryID
        public long VRD_RefID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public int VendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public string DisplayMaker => !string.IsNullOrEmpty(VendID1)
            ? $"{VenderName} ({VendID1})"
            : VenderName;

        public int Repair_RefID { get; set; }
        public string RepairType { get; set; } = string.Empty;

        public decimal Qty { get; set; }
        public decimal IssQty { get; set; }
        public decimal BalanceQty => Qty - IssQty;
        public decimal Rate { get; set; }

        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string ProcessDescription { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
    }

    public class ReWorkStagedItem
    {
        public long EntryID { get; set; }          // VendRcvdDetailReWorkDetail.EntryID
        public string LotNo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public int VendID { get; set; }
        public string DisplayMaker { get; set; } = string.Empty;
        public int Repair_RefID { get; set; }
        public string RepairType { get; set; } = string.Empty;
        public decimal Qty { get; set; }
        public decimal Rate { get; set; }
    }

    public class ReWorkIssuanceFilter
    {
        public string ArticleId { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public int RepairTypeId { get; set; } = 0; // 0 = <All>
        public int DateRangeIndex { get; set; } = 0; // 0=Today, 1=Last 7d, 2=Last 15d, 3=Last 30d, 4=Custom
        public DateTime DtFrom { get; set; } = DateTime.Today;
        public DateTime DtTo { get; set; } = DateTime.Today;
    }

    public class SaveReWorkIssuanceRequest
    {
        public string ItemID { get; set; } = string.Empty;
        public int MakerID { get; set; }
        public List<ReWorkStagedItem> StagedItems { get; set; } = new List<ReWorkStagedItem>();
        public bool PrintPTC { get; set; } = false;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }

    public class ReWorkIssuanceResult
    {
        public bool Success { get; set; }
        public string NewLotNo { get; set; } = string.Empty;
        public string MakerIssNo { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }
}
