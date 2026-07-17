using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class TotalDemandModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public double OrderAmt { get; set; }
        public string OrderType { get; set; } = string.Empty;
    }
}
