using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class SFFromLocationModel
    {
        public int EntryID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public int Store_RefID { get; set; }
        public int Shelf_RefID { get; set; }
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public string StoreName { get; set; } = string.Empty;
        public string LotNo_Manual { get; set; } = string.Empty;
        public int AvailableQty { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public string DisplayText => $"{RackNo} {{{ShelfNo}}} {{{AvailableQty}}} ({LotNo_Manual})";
    }

    public class SFTargetStoreModel
    {
        public int EntryID { get; set; }
        public string StoreName { get; set; } = string.Empty;
    }

    public class SFTargetRackModel
    {
        public string RackNo { get; set; } = string.Empty;
    }

    public class SFTargetShelfModel
    {
        public int EntryID { get; set; }
        public int Store_RefID { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public string LocationDisplay => $"{RackNo} {{{ShelfNo}}}";
    }

    public class SFMovementRequest
    {
        public int SOO_EntryID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public int TargetShelfID { get; set; }
        public string TargetLocationDisplay { get; set; } = string.Empty;
        public int QtyToMove { get; set; }
        public string UserName { get; set; } = string.Empty;
    }
}
