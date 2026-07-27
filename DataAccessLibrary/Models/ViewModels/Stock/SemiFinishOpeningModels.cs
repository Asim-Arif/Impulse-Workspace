using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class SemiFinishOpeningModel
    {
        public int EntryID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty; // Process Description
        public int Qty { get; set; }
        public int IssdQty { get; set; }
        public int AvailableQty => Qty - IssdQty;
        public string Location { get; set; } = string.Empty;
        public string StoreName { get; set; } = string.Empty;
        public int Store_RefID { get; set; }
        public int Shelf_RefID { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public DateTime? DTEntry { get; set; }
        public string LotNo_Manual { get; set; } = string.Empty;
        public string BatchNo_Manual { get; set; } = string.Empty;
        public string Mill_Certificate_No { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }

    public class ItemProcessViewModel
    {
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
        public int ItemSNo { get; set; }
    }

    public class StoreLookupModel
    {
        public int EntryID { get; set; }
        public string StoreName { get; set; } = string.Empty;
    }

    public class StoreShelfLookupModel
    {
        public int EntryID { get; set; }
        public int Store_RefID { get; set; }
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public string StoreName { get; set; } = string.Empty;
        public string LocationDisplay => $"{RackNo} {{{ShelfNo}}}";
    }

    public class CreateSemiFinishOpeningRequest
    {
        public string ItemID { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public int Qty { get; set; }
        public int Shelf_RefID { get; set; }
        public string Location { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public string LotNo_Manual { get; set; } = string.Empty;
        public string BatchNo_Manual { get; set; } = string.Empty;
        public string Mill_Certificate_No { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }
}
