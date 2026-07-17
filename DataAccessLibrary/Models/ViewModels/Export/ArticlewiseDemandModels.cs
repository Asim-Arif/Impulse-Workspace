using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class ArticlewiseDemandModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string CompItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public double Qty { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
    }
}
