using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class PostPrcModel
    {
        public int EntryID { get; set; }
        public string CustomInvoice { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Currency { get; set; } = string.Empty;
        public decimal ExchRate { get; set; }
        public decimal AmountRealized { get; set; }
        public decimal AmountInRs => Math.Round(AmountRealized * ExchRate, 2);
        
        public string BillNo { get; set; } = string.Empty;
        public DateTime BillDate { get; set; } = DateTime.Today;
        public string SerialNo { get; set; } = string.Empty;
        public DateTime SerialDate { get; set; } = DateTime.Today;
        
        public DateTime RealizationDate { get; set; } = DateTime.Today;
        public DateTime PostingDate { get; set; } = DateTime.Today;
        
        public string BankAccNo { get; set; } = string.Empty;
        
        public List<PrcDeductionModel> Deductions { get; set; } = new();
    }

    public class PrcDeductionModel
    {
        public string Title { get; set; } = string.Empty;
        public string AccountNo { get; set; } = string.Empty;
        public decimal Amount { get; set; }
        public string Currency { get; set; } = "PKR";
    }
}
