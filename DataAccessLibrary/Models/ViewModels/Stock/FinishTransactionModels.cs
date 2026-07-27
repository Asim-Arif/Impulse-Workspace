using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class FinishTransactionModel
    {
        public string EntryNo { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string RcvdBy { get; set; } = string.Empty;
        public string RcvdFrom { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public double Qty { get; set; }
        public string Location { get; set; } = string.Empty;
        public int ChildEntryID { get; set; }
        public int EntryType { get; set; } // 1 = Receiving, 0 = Issuance
        public int? Shelf_RefID { get; set; }
        public int? Store_RefID { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public string LotNo_Manual { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;

        public double SignedQty => EntryType == 1 ? Qty : -Qty;
    }

    public class FinishTransactionFilterModel
    {
        public DateTime FromDate { get; set; } = DateTime.Today;
        public DateTime ToDate { get; set; } = DateTime.Today;
        public bool IncludeIssuance { get; set; } = true;
        public bool IncludeReceiving { get; set; } = true;
        public int StoreRefId { get; set; } = 0;
        public string ItemID { get; set; } = "0";
    }
}
