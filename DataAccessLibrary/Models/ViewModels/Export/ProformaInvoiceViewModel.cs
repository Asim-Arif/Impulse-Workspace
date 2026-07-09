using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    /// <summary>Master DTO for FPInvoice header</summary>
    public class ProformaInvoiceViewModel
    {
        public int PInvoiceID { get; set; }
        public string PInvoice { get; set; } = string.Empty;
        public DateTime DT { get; set; } = DateTime.Today;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Currency { get; set; } = string.Empty;

        // Header shipping fields
        public string SMethod { get; set; } = string.Empty;
        public string TradeTerms { get; set; } = string.Empty;
        public string PaymentTerms { get; set; } = string.Empty;
        public string LCNo { get; set; } = string.Empty;
        public string PartialShipment { get; set; } = "Allowed";
        public string TransShipment { get; set; } = "Allowed";
        public string Packaging { get; set; } = string.Empty;
        public DateTime DeliveryDT { get; set; } = DateTime.Today.AddMonths(2);

        // Port IDs
        public int? LoadPort { get; set; }
        public int? DischargePort { get; set; }

        // Banking (Other Info tab)
        public int? BankID { get; set; }           // Customer bank (Through)
        public int? OurBankID { get; set; }         // Company bank account

        // Agents
        public int? ShippingAgentID { get; set; }
        public int? InsuranceAgentID { get; set; }

        // Company
        public int CompanyRefID { get; set; }

        // Weights & packages
        public string GrossWeight { get; set; } = string.Empty;
        public string NetWeight { get; set; } = string.Empty;
        public string Packages { get; set; } = string.Empty;
        public double SteelWeight { get; set; }

        // Text fields
        public string InvoiceOf { get; set; } = "Surgical Instruments";
        public string DeliveryDetail { get; set; } = string.Empty;
        public string ShippingAddress { get; set; } = string.Empty;
        public string InvoiceTo { get; set; } = string.Empty;
        public string Declarations { get; set; } = string.Empty;
        public string BatchNo { get; set; } = string.Empty;

        // Total / discount section
        public string TotalDescription { get; set; } = "TOTAL";
        public double DiscountPer { get; set; }
        public double DiscountAmount { get; set; }
        public string DiscountDescription { get; set; } = "DISCOUNT";
        public double InsuranceFreightAmount { get; set; }
        public string InsuranceDescription { get; set; } = "FREIGHT";
        public double InsuranceFreightAmountII { get; set; }
        public string InsuranceDescriptionII { get; set; } = string.Empty;
        public string CIFDescription { get; set; } = "GRAND TOTAL";

        // Notes
        public string Notes1 { get; set; } = string.Empty;
        public string Notes2 { get; set; } = string.Empty;

        // Flags
        public bool PrintOnSave { get; set; }

        // Derived / UI only
        public double TotalAmount { get; set; }
        public double GrandTotal { get; set; }
        public string AmountInWords { get; set; } = string.Empty;

        // Child collections
        public List<ProformaLineItemModel> LineItems { get; set; } = new();
        public List<ProformaOtherChargeModel> OtherCharges { get; set; } = new();

        // Deleted IDs (for save)
        public List<int> DeletedLineItemIds { get; set; } = new();
    }

    /// <summary>A row in the Proforma Details grid (FProformaOrders + view data)</summary>
    public class ProformaLineItemModel
    {
        public int EntryID { get; set; }            // FProformaOrders.EntryID (0 = new)
        public int OrderItemID { get; set; }         // FK to FOrderItems.ID
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string CompItemCode { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public double Price { get; set; }
        public double CustomPrice { get; set; }
        public int Qty { get; set; }
        public int InvQty { get; set; }
        public int InvoicedQty { get; set; }        // Already invoiced (for min qty constraint)
        public bool IsHighlighted { get; set; }

        // Sub-detail row (delivery / quality / stamps)
        public DateTime? DeliveryDTItem { get; set; }
        public string QualityItem { get; set; } = string.Empty;
        public string StampsItem { get; set; } = string.Empty;

        // Computed
        public double Amount => Math.Round(Price * Qty, 4);
    }

    /// <summary>An order item available to add to proforma (VFCustOrders row)</summary>
    public class OrderItemForProformaModel
    {
        public int OrderItemID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string CompItemCode { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public double Price { get; set; }
        public double CustomPrice { get; set; }
        public int InvQty { get; set; }             // Available qty to add
        public string TradeTerms { get; set; } = string.Empty;
        public string PartialShipment { get; set; } = string.Empty;
        public string TransShipment { get; set; } = string.Empty;
        public string PaymentTerms { get; set; } = string.Empty;
        public string Packaging { get; set; } = string.Empty;
        public DateTime DeliveryDT { get; set; }
        public DateTime DT { get; set; }
        public DateTime? DeliveryDTItem { get; set; }
        public string QualityItem { get; set; } = string.Empty;
        public string StampsItem { get; set; } = string.Empty;

        // Computed
        public double Amount => Math.Round(Price * QtyToAdd, 4); // Use QtyToAdd for amount calculation
        // Qty input for adding to proforma
        public int QtyToAdd { get; set; }
        public bool IsSelected { get; set; }
    }

    /// <summary>A row in the Other Charges grid (FPInvoiceOtherCharges)</summary>
    public class ProformaOtherChargeModel
    {
        public int SNo { get; set; }
        public string ChargesDetail { get; set; } = string.Empty;
        public double ChargesAmt { get; set; }
    }

    /// <summary>Customer bank (FCustBanks)</summary>
    public class CustomerBankModel
    {
        public int BankID { get; set; }
        public string BankName { get; set; } = string.Empty;
        public string Address1 { get; set; } = string.Empty;
    }

    /// <summary>Company bank account (BankList)</summary>
    public class CompanyBankAccountModel
    {
        public int BankID { get; set; }
        public string BankName { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
    }

    /// <summary>Port of Loading / Discharge (Ports table)</summary>
    public class PortModel
    {
        public int PortID { get; set; }
        public string Port { get; set; } = string.Empty;
        public int CityID { get; set; }
    }

    /// <summary>City (Cities table)</summary>
    public class CityModel
    {
        public int CityID { get; set; }
        public string CityName { get; set; } = string.Empty;
    }

    /// <summary>Shipping/Insurance Agent</summary>
    public class AgentLookupModel
    {
        public int EntryID { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    /// <summary>Shipping address preset (ForeignCustomersShippingAddresses)</summary>
    public class ShippingAddressPresetModel
    {
        public int EntryID { get; set; }
        public string ShippingAddress { get; set; } = string.Empty;
    }

    /// <summary>Invoice To preset (ForeignCustomers_InvoiceTo)</summary>
    public class InvoiceToPresetModel
    {
        public int EntryID { get; set; }
        public string InvoiceTo { get; set; } = string.Empty;
    }
}
