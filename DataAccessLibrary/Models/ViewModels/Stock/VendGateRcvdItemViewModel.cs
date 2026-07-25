using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendGateRcvdItemViewModel
    {
        public string OrderNo { get; set; }
        public string MaterialID { get; set; }
        public string MaterialName { get; set; }
        public double? QtyOrdered { get; set; }
        public double? QtyToRcv { get; set; }
        public double? QtyRcvd { get; set; }
        public string StoreName { get; set; }
        public float? Rate { get; set; }
        public int? Sheets { get; set; }
        public int? WORef { get; set; }
        public byte? IssuanceType { get; set; }
        
        // For UI purposes
        public string Unit { get; set; }
        public bool IsEditing { get; set; }
    }
}
