using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CustomerQuotationItemViewModel
    {
        public int EntryID { get; set; } = 0;
        public int QuotationNo { get; set; } = 0;
        public string ItemCode { get; set; } = string.Empty;
        public int Qty { get; set; } = 0;
        public decimal Price { get; set; } = 0;
        public DateTime? DeliveryDT { get; set; }
        public string QualityType { get; set; } = string.Empty;
        public string ItemRemarks { get; set; } = string.Empty;

        // Custom fields for UI Binding
        public string ItemName { get; set; } = string.Empty;
    }
}
