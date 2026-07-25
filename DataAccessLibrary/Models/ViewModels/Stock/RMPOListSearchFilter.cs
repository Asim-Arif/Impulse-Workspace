using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class RMPOListSearchFilter
    {
        public string VendorID { get; set; }
        public string MaterialID { get; set; }
        public int? RMGroupID { get; set; }
        public byte? POType { get; set; }
        public int DateRangeIndex { get; set; }
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public string OrderNo { get; set; }
        public bool ShowOpenPOsOnly { get; set; }
        public bool ShowLateOrdersOnly { get; set; }
        public bool ShowSamplePOs { get; set; }
    }
}
