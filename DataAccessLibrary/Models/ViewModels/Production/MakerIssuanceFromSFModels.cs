using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class SFLotLookupResultModel
    {
        public long EntryID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public decimal RcvdQty { get; set; }
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public long Shelf_RefID { get; set; }
        public long SOO_RefID { get; set; }
        public decimal SplitQty { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }

    public class SFStockItemLookupModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string DisplayText => $"[{ItemID}] {ItemName}";
    }

    public class SFStockOpeningLocationModel
    {
        public long EntryID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public long Shelf_RefID { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public decimal InHandQty { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public string DisplayText => $"{StoreName} - Rack {RackNo} / Shelf {ShelfNo} (In Hand: {InHandQty:N0}, Lot: {LotNo})";
    }

    public class UnshippedOrderLookupModel
    {
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string DisplayText => string.IsNullOrWhiteSpace(InternalRefNo) ? OrderNo : $"{OrderNo} ({InternalRefNo})";
    }

    public class StagedSFIssuanceItemModel
    {
        public int LineNo { get; set; }
        public long OpeningEntryID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public decimal IssQty { get; set; }
        public decimal Rate { get; set; }
        public decimal RateDozen => Rate * 12m;
        public decimal Value => IssQty * Rate;
        public DateTime ReturnDT { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public int ReturnProcessID { get; set; }
        public string ReturnProcessName { get; set; } = string.Empty;
    }

    public class CreateSFIssuanceHeaderModel
    {
        public long VendID { get; set; }
        public DateTime DT { get; set; } = DateTime.Today;
        public int ProcessID { get; set; }
        public int ReturnProcessID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string SpecialInstructions { get; set; } = string.Empty;
        public decimal ExcessQtyPercentage { get; set; } = 0m;
        public int MaximumRcvingsAgainstPO { get; set; } = 1;
        public string IssEmpID { get; set; } = string.Empty;
        public string CheckerEmpID { get; set; } = string.Empty;
        public bool IssuanceByLotNo { get; set; } = false;
        public string SearchedLotNo { get; set; } = string.Empty;
    }

    public class CreateSFIssuanceLineModel
    {
        public long OpeningEntryID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public decimal Rate { get; set; }
        public decimal IssQty { get; set; }
        public int ReturnProcessID { get; set; }
        public DateTime ReturnDT { get; set; }
        public int Priority { get; set; } = 1;
        public string LotNo { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
    }
}
