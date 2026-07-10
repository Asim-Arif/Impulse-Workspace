using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class ReceivePaymentModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Currency { get; set; } = string.Empty;
        public string CustomInvoice { get; set; } = string.Empty;
        public DateTime Date { get; set; } = DateTime.Today;
        public decimal Amount { get; set; }
        public decimal BankCharges { get; set; }
        public decimal ExchangeRate { get; set; }
        public string Description { get; set; } = string.Empty;
        public int BankID { get; set; }
        public string FDBCNo { get; set; } = string.Empty;
        public bool IsCustomPayment { get; set; }
    }
}
