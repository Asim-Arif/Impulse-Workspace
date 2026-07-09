using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class AdvancePaymentUsageViewModel
    {
        public int EntryID { get; set; }
        public int RefID { get; set; }
        public string CustomInvoice { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }
}
