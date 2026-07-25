using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class PlacedMaterialItemViewModel
    {
        public int EntryID { get; set; }
        public DateTime RcvDate { get; set; }
        public string VendorName { get; set; }
        public string PORefNo { get; set; }
        public string MaterialName { get; set; }
        public string StoreName { get; set; }
        public string RackNo { get; set; }
        public string ShelfNo { get; set; }
        public decimal QtyPlaced { get; set; }
        public string Unit { get; set; }
        public string RMID1 { get; set; }
        
        // Derived property for Location
        public string Location => $"{StoreName} {RackNo} {ShelfNo}".Trim();
    }
}
