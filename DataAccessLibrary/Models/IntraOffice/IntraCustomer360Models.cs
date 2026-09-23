using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class Customer360Dto
    {
        public int Id { get; set; }
        public string CustomerCode { get; set; } = string.Empty;
        public string CompanyName { get; set; } = string.Empty;
        public string CustomerType { get; set; } = "Wholesaler"; // Wholesaler, Distributor, OEM, Retailer, Brand
        public string Country { get; set; } = string.Empty;
        public string? City { get; set; }
        public string? BillingAddress { get; set; }
        public string? ShippingAddress { get; set; }
        public string? PrimaryEmail { get; set; }
        public string? PrimaryPhone { get; set; }
        public string? WhatsApp { get; set; }
        public string? Website { get; set; }
        public string? TaxVatEoriNumber { get; set; }
        public string PaymentTerms { get; set; } = "Letter of Credit (LC at sight)";
        public decimal CreditLimit { get; set; } = 250000;
        public string PreferredCurrency { get; set; } = "USD";
        public string Incoterm { get; set; } = "FOB";
        public string Status { get; set; } = "Active"; // Active, Inactive, OnHold
        public string Rating { get; set; } = "Gold"; // Platinum, Gold, Silver, Bronze
        public string? Notes { get; set; }
        public string? AssignedRep { get; set; } = "Zeeshan";
        public bool IsConvertedToForeignCustomer { get; set; } = false;
        public string? ForeignCustomerCode { get; set; }
        public int LeadId { get; set; }
        public string? LeadStatus { get; set; } = "Qualified";
        public string? Priority { get; set; } = "Medium";

        // Financial KPIs
        public decimal TotalSalesRevenue { get; set; } = 0;
        public decimal OutstandingArBalance { get; set; } = 0;

        // CRM Collections
        public List<CustomerContactDto> Contacts { get; set; } = new();
        public List<CustomerInquiryDto> Inquiries { get; set; } = new();
        public List<CustomerActivityDto> Activities { get; set; } = new();

        // Impulse Shell Counts
        public int QuotationsCount { get; set; } = 1;
        public int OrdersCount { get; set; } = 0;
        public int InvoicesCount { get; set; } = 0;
        public int ComplaintsCount { get; set; } = 0;
        public int SamplesCount { get; set; } = 0;
        public int ComplianceCount { get; set; } = 0;
    }

    public class CustomerContactDto
    {
        public int Id { get; set; }
        public int CustomerId { get; set; }
        public string Name { get; set; } = string.Empty;
        public string? Designation { get; set; }
        public string? Email { get; set; }
        public string? Phone { get; set; }
        public string? WhatsApp { get; set; }
        public bool IsPrimary { get; set; } = false;
    }

    public class CustomerInquiryDto
    {
        public int Id { get; set; }
        public string InquiryNumber { get; set; } = string.Empty;
        public DateTime InquiryDate { get; set; } = DateTime.UtcNow;
        public decimal? TargetPrice { get; set; }
        public string? DestinationPort { get; set; }
        public int ItemsCount { get; set; } = 1;
        public string Status { get; set; } = "Active";
        public string? ProductSummary { get; set; }
    }

    public class CustomerActivityDto
    {
        public int Id { get; set; }
        public int CustomerId { get; set; }
        public string ActivityType { get; set; } = "Call"; // Call, WhatsApp, Email, Meeting, InternalNote
        public string Subject { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string PerformedBy { get; set; } = "System";
        public DateTime ActivityDate { get; set; } = DateTime.UtcNow;
    }

    public class Customer360LookupDto
    {
        public int LeadId { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string Status { get; set; } = "New";
        public bool IsConverted { get; set; }
        public string? ConvertedCustCode { get; set; }
        public bool IsForeignCustomer { get; set; }
    }

    public class ConvertCustomerToForeignCustomerModel
    {
        public int LeadId { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string? City { get; set; }
        public string? Address { get; set; }
        public string? Phone1 { get; set; }
        public string? Email1 { get; set; }
        public string? URL { get; set; }
        public string? Curr { get; set; } = "USD";
        public string? TradeTerms { get; set; } = "FOB";
        public string? PaymentTerms { get; set; } = "Letter of Credit (LC at sight)";
        public string? AccNo { get; set; }
        public string? Cont1name { get; set; }
        public string? Cont1Designation { get; set; }
        public string? Cont1Email { get; set; }
        public string? cont1Phone { get; set; }
        public string? Cont1Mobile { get; set; }
        public string? SpecialInstructions { get; set; }
        public string? Customer_Type { get; set; } = "Wholesaler";
        public string ConvertedBy { get; set; } = string.Empty;
    }
}
