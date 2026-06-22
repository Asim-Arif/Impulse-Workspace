using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    // Main foreign customer model; field names match VB6 Tag values from frmFCustomer.frm
    public class ForeignCustomerViewModel
    {
        // Header / identity
        public string CustCode { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Curr { get; set; } = string.Empty;

        // Company Info tab
        public string Phone1 { get; set; } = string.Empty;
        public string Phone2 { get; set; } = string.Empty;
        public string Phone3 { get; set; } = string.Empty;
        public string Fax1 { get; set; } = string.Empty;
        public string Fax2 { get; set; } = string.Empty;
        public string Fax3 { get; set; } = string.Empty;
        public string EMail1 { get; set; } = string.Empty;
        public string EMail2 { get; set; } = string.Empty;
        public string URL { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string FakeAddress { get; set; } = string.Empty;
        public string PaymentTerms { get; set; } = string.Empty;
        public string TradeTerms { get; set; } = string.Empty;
        public string DTFormat { get; set; } = "dd-MMM-yyyy";
        public string FinishingQuality { get; set; } = string.Empty;
        public string DefaultShipMethod { get; set; } = string.Empty;
        public string CustomerSource { get; set; } = string.Empty;
        public int? DefaultPort { get; set; }
        public string OuterPackingLabel { get; set; } = string.Empty;
        public string InnerPackingLabel { get; set; } = string.Empty;
        public string Inner_Label_Manual_I { get; set; } = string.Empty;
        public string Inner_Label_Manual_II { get; set; } = string.Empty;
        public string Stamps { get; set; } = string.Empty;
        public string RcvdVia { get; set; } = string.Empty;
        public string SpecialInstructions { get; set; } = string.Empty;
        public string FC_Note_I { get; set; } = string.Empty;
        public string FC_Note_II { get; set; } = string.Empty;
        public decimal PaymentDays { get; set; }
        public decimal MaxDiscount { get; set; }
        public bool AcceptsExtraQty { get; set; }
        public bool LateOrderAlerts { get; set; }
        public bool ShowCustomerRef { get; set; }
        public bool Active { get; set; } = true;

        // Contact Info tab — Contact Person 1
        public string Cont1Name { get; set; } = string.Empty;
        public string Cont1Designation { get; set; } = string.Empty;
        public string Cont1Email { get; set; } = string.Empty;
        public string cont1Phone { get; set; } = string.Empty;
        public string Cont1Mobile { get; set; } = string.Empty;
        public string Cont1Skype { get; set; } = string.Empty;

        // Contact Info tab — Contact Person 2
        public string Cont2Name { get; set; } = string.Empty;
        public string Cont2Designation { get; set; } = string.Empty;
        public string Cont2Email { get; set; } = string.Empty;
        public string Cont2Phone { get; set; } = string.Empty;
        public string Cont2Mobile { get; set; } = string.Empty;
        public string Cont2Skype { get; set; } = string.Empty;

        // Contact Info — Misc. Information
        public string NearestAirport { get; set; } = string.Empty;
        public string NearestRailwaystation { get; set; } = string.Empty;
        public string FedexNo { get; set; } = string.Empty;
        public string DHLNo { get; set; } = string.Empty;

        // Agent Info tab
        public string AgentName { get; set; } = string.Empty;
        public string AgentEmail { get; set; } = string.Empty;
        public string AgentPhone1 { get; set; } = string.Empty;
        public string AgentPhone2 { get; set; } = string.Empty;
        public string AgentFax1 { get; set; } = string.Empty;
        public string AgentFax2 { get; set; } = string.Empty;
        public string AgentAddress { get; set; } = string.Empty;
        public string AgentURL { get; set; } = string.Empty;

        // Agent — Contact Person
        public string ACName { get; set; } = string.Empty;
        public string ACEmail { get; set; } = string.Empty;
        public string ACPhone { get; set; } = string.Empty;
        public string ACMobile { get; set; } = string.Empty;

        // Accounts Info tab (mostly read-only)
        public string AccNo { get; set; } = string.Empty;

        // Lists — Bank Info tab
        public List<ForeignCustomerBankViewModel> Banks { get; set; } = new();

        // Lists — Shipping Addresses tab
        public List<ForeignCustomerAddressViewModel> ShippingAddresses { get; set; } = new();

        // Lists — Events tab
        public List<ForeignCustomerEventViewModel> Events { get; set; } = new();

        // Lists — Invoice To tab
        public List<ForeignCustomerInvoiceToViewModel> InvoiceTo { get; set; } = new();
    }

    public class ForeignCustomerBankViewModel
    {
        public int BankID { get; set; }
        public string BankName { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Phone1 { get; set; } = string.Empty;
        public string Phone2 { get; set; } = string.Empty;
        public string Fax1 { get; set; } = string.Empty;
        public string Fax2 { get; set; } = string.Empty;
        public string Email1 { get; set; } = string.Empty;
        public string Email2 { get; set; } = string.Empty;
        public string Address1 { get; set; } = string.Empty;
        public string Address2 { get; set; } = string.Empty;
        public bool DefaultBank { get; set; }
    }

    public class ForeignCustomerAddressViewModel
    {
        public int EntryID { get; set; }
        public string ShippingAddress { get; set; } = string.Empty;
        public bool IsEditing { get; set; }
    }

    public class ForeignCustomerEventViewModel
    {
        public int EntryID { get; set; }
        public string Description { get; set; } = string.Empty;
        public DateTime EventDT { get; set; } = DateTime.Today;
        public int RepeatedType { get; set; }
        public string RepeatDescription { get; set; } = string.Empty;
    }

    public class ForeignCustomerInvoiceToViewModel
    {
        public int EntryID { get; set; }
        public string InvoiceTo { get; set; } = string.Empty;
        public bool IsEditing { get; set; }
    }

    // Lightweight model for the customers list page
    public class ForeignCustomerListViewModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Curr { get; set; } = string.Empty;
        public bool Active { get; set; }
    }
}
