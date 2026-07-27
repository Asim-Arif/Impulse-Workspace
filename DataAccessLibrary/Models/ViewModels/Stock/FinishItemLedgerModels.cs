using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class ItemHeaderInfoModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public string ItemSize { get; set; } = string.Empty;
        public string SizeUnit { get; set; } = string.Empty;
        public string ItemColor { get; set; } = string.Empty;
        public double TotalStock { get; set; }
    }

    public class FinishItemLedgerRowModel
    {
        public DateTime? DT { get; set; }
        public string RcvNo { get; set; } = string.Empty;
        public string RcvdBy { get; set; } = string.Empty;
        public string RcvdFrom { get; set; } = string.Empty;
        public double QtyIn { get; set; }
        public double QtyOut { get; set; }
        public int EntryType { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string LotNo_Manual { get; set; } = string.Empty;
        public string BatchNo_Manual { get; set; } = string.Empty;
        public int? Shelf_RefID { get; set; }
        public string Location { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public double OpeningBalance { get; set; }
        public double RunningBalance { get; set; }
    }

    public class FinishItemLedgerFilterModel
    {
        public string ItemID { get; set; } = string.Empty;
        public DateTime FromDate { get; set; } = DateTime.Today.AddMonths(-1);
        public DateTime ToDate { get; set; } = DateTime.Today;
    }
}
