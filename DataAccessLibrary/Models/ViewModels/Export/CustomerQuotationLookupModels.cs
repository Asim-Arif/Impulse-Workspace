using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CustomerQuotationListItemModel
    {
        public int QuotationNo { get; set; }
        public DateTime DT { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string TradeTerms { get; set; } = string.Empty;
        public int CompanyRefID { get; set; }
        public string Curr { get; set; } = string.Empty;
        public string Validity { get; set; } = string.Empty;
        public string PaymentTerms { get; set; } = string.Empty;
        public DateTime? DeliveryDT { get; set; }
        public string ShipVia { get; set; } = string.Empty;
        public string Packing { get; set; } = string.Empty;
        public string DeliveryTime { get; set; } = string.Empty;
        public int? BankID { get; set; }
        
        // Custom field for UI display (joined with ForeignCustomers)
        public string CustomerName { get; set; } = string.Empty;
    }
}
