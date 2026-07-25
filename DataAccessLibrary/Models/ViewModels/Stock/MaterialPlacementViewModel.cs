using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class MaterialPlacementListViewModel
    {
        public int EntryID { get; set; } // VMaterialsForPlacement.EntryID
        public string RcvID { get; set; }
        public string MaterialID { get; set; }
        public string MaterialName { get; set; }
        public string VendorName { get; set; }
        public decimal QtyPassed { get; set; }
        public decimal QtyPlaced { get; set; }
        public decimal Balance { get; set; }
        public string Unit { get; set; }
        public bool IsCCItem { get; set; } // CCItem
        public int RISD_EntryID { get; set; } // RISD_EntryID
        public DateTime RcvDate { get; set; }
    }

    public class StoreViewModel
    {
        public int EntryID { get; set; }
        public string StoreName { get; set; }
    }

    public class RackViewModel
    {
        public int EntryID { get; set; }
        public string RackNo { get; set; }
    }

    public class ShelfViewModel
    {
        public int EntryID { get; set; }
        public string ShelfNo { get; set; }
    }

    public class PlacementSaveRequest
    {
        public int Rcvd_RefID { get; set; }
        public int Shelf_RefID { get; set; }
        public decimal QtyPlaced { get; set; }
        public string BatchNo { get; set; }
        public string LotNo { get; set; }
        
        // Context properties
        public bool IsCCItem { get; set; }
        public int RISD_RefID { get; set; }
    }
}
