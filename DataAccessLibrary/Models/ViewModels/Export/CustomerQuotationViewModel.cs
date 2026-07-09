using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CustomerQuotationViewModel
    {
        public int QuotationNo { get; set; } = 0;
        public DateTime DT { get; set; } = DateTime.Today;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Validity { get; set; } = string.Empty;
        public string PaymentTerms { get; set; } = string.Empty;
        public DateTime? DeliveryDT { get; set; }
        public string ShipVia { get; set; } = string.Empty;
        public string Packing { get; set; } = string.Empty;
        public string DeliveryTime { get; set; } = string.Empty;
        public int CompanyRefID { get; set; } = 1;
        public string TradeTerms { get; set; } = string.Empty;
        public int? BankID { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
        public DateTime? DTEntry { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public string CustomerRef { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public string AttentionPerson { get; set; } = string.Empty;
        public string SwiftCode { get; set; } = string.Empty;
        public string GrossWeight { get; set; } = string.Empty;
        public string NetWeight { get; set; } = string.Empty;
        public int? NoOfPackages { get; set; }
        public string RevisionHistory { get; set; } = string.Empty;
        public string Destination { get; set; } = string.Empty;

        // Custom field for UI binding to display currency from foreign customer
        public string Currency { get; set; } = string.Empty;

        public List<CustomerQuotationItemViewModel> QuotationItems { get; set; } = new();
        public List<CustomerQuotationOtherChargeViewModel> OtherCharges { get; set; } = new();
    }
}
