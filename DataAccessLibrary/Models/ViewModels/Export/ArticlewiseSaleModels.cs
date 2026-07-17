using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class ArticlewiseSaleModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string CustomInvoice { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public double Qty { get; set; }
    }
}
