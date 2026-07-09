using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CustomerQuotationOtherChargeViewModel
    {
        public int EntryID { get; set; } = 0;
        public int QuotationNo { get; set; } = 0;
        public string ChargesDetail { get; set; } = string.Empty;
        public decimal ChargesAmt { get; set; } = 0;
        public int SNo { get; set; } = 0;
    }
}
