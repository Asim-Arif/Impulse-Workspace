using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class AdvancePaymentViewModel
    {
        public int EntryID { get; set; }
        public DateTime? DT { get; set; } = DateTime.Today;
        public string AdviceNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public decimal Amount { get; set; }
        public string ImporterAddress { get; set; } = string.Empty;
        public string ExportOf { get; set; } = "SURGICAL INSTRUMENTS";
        public decimal ExchRate { get; set; }
        public string VchrNo { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
        public DateTime? DTEntry { get; set; }
        public string BankAccNo { get; set; } = string.Empty;
        public int? CustID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
    }

    public class AdvancePaymentListViewModel : AdvancePaymentViewModel
    {
        // Joined / Derived fields for the UI List
        public string Status 
        { 
            get 
            {
                if (string.IsNullOrEmpty(VchrNo))
                {
                    return ExchRate <= 0 ? "Update Exchange Rate" : "Ready to Post";
                }
                return "Posted to Financial";
            }
        }
        
        public string CustomerName { get; set; } = string.Empty;
        public string BankName { get; set; } = string.Empty;
    }

    public class AdvancePaymentBankModel
    {
        public string BankAccNo { get; set; } = string.Empty;
        public string BankName { get; set; } = string.Empty;
    }
}
