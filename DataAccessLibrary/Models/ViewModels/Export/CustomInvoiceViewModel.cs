using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    /// <summary>Master DTO for CustomInvoice header</summary>
    public class CustomInvoiceViewModel
    {
        public string CustomInvoiceNo { get; set; } = string.Empty; // PK CustomInvoice
        public DateTime DT { get; set; } = DateTime.Today;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Currency { get; set; } = string.Empty; // CustomInvoice.CustCurrency

        // Header fields
        public string ShippingMode { get; set; } = string.Empty;
        public string TradeTerms { get; set; } = string.Empty;
        public string PaymentTerms { get; set; } = string.Empty;
        public string LCNo { get; set; } = string.Empty;
        public DateTime? LCDate { get; set; }
        public string PartialShipment { get; set; } = "Allowed";
        public string TransShipment { get; set; } = "Allowed";
        public string Packaging { get; set; } = string.Empty;

        // Flags
        public bool NoSRO { get; set; }
        public bool CIFTradeTerms { get; set; }
        public bool SampleInvoice { get; set; }
        public bool FreeSamples { get; set; }

        // GD & Custom dates
        public string GDNo { get; set; } = string.Empty;
        public DateTime? CustomDT { get; set; }

        // Ports
        public int? LoadPortID { get; set; }
        public int? DisPortID { get; set; }

        // Banking & Agents
        public int? OurBankID { get; set; }
        public int? ShippingAgentID { get; set; }
        public int? InsuranceAgentID { get; set; }

        // Company
        public int CompanyRefID { get; set; }

        // Other Strings
        public string AWBNo { get; set; } = string.Empty;
        public DateTime? AWBNoDT { get; set; }
        public string FormE { get; set; } = string.Empty;
        public DateTime? FormEDT { get; set; }
        public string Vessel { get; set; } = string.Empty;
        public double GrossWeight { get; set; }

        public string DeclarationsBalls { get; set; } = string.Empty;
        public string DeclarationsGarments { get; set; } = string.Empty;
        public string DeclarationsGloves { get; set; } = string.Empty;
        public string Consignee { get; set; } = string.Empty;

        public string ComDeclarationsBalls { get; set; } = string.Empty;
        public string ComDeclarationsGarments { get; set; } = string.Empty;
        public string ComDeclarationsGloves { get; set; } = string.Empty;
        public string ComNotify { get; set; } = string.Empty;
        public string ComConsignee { get; set; } = string.Empty;
        public string ComPayTerms { get; set; } = string.Empty;
        public string ComAgainst { get; set; } = string.Empty; // Maps to ShippingAddress
        public string ComDrawnUnder { get; set; } = string.Empty;
        public string ComSpecial { get; set; } = string.Empty;
        
        public string ManDischarge { get; set; } = string.Empty;
        public string ManDeclaration { get; set; } = string.Empty;
        public string Warehouse { get; set; } = string.Empty;
        public string Gate { get; set; } = string.Empty;
        public string ContainerNo { get; set; } = string.Empty;
        public string SealNo { get; set; } = string.Empty;
        public DateTime? DTETA { get; set; }
        
        public string SRONo { get; set; } = string.Empty;
        public string InvoiceOf { get; set; } = "Surgical Instruments";
        public string GatePassNo { get; set; } = string.Empty;
        public DateTime? GatePassDT { get; set; }
        public string DTREDescription { get; set; } = string.Empty;
        public string InvoiceDescription { get; set; } = string.Empty;
        
        public string InvoiceTo { get; set; } = string.Empty;
        public int TotalCartons { get; set; }


        // UI only / Derived
        public double TotalAmount { get; set; }
        public double GrandTotal { get; set; }
        public string AmountInWords { get; set; } = string.Empty;

        // Child collections
        public List<CustomInvoiceItemModel> LineItems { get; set; } = new();
        public List<CustomPListModel> PackingListItems { get; set; } = new();
        public List<CustomInvoiceOtherChargeModel> OtherCharges { get; set; } = new();
        public List<AdvancePaymentModel> AdvancePayments { get; set; } = new();

        // Track deletions
        public List<int> DeletedLineItemIds { get; set; } = new();
        public List<int> DeletedPackingListIds { get; set; } = new();
        public List<int> DeletedOtherChargeIds { get; set; } = new();
    }

    public class CustomInvoiceItemModel
    {
        public int EntryID { get; set; }
        public int RefID { get; set; } // FK to FProformaOrders.EntryID
        public string ProformaInvoice { get; set; } = string.Empty; // Added to trace back
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public int Qty { get; set; }
        public double Price { get; set; }
        public double CustomPrice { get; set; }
        public double RebateRate { get; set; }
        public string DTRENo { get; set; } = string.Empty;
        public bool ItemForLabTest { get; set; }
        public string BatchNo { get; set; } = string.Empty;
        public string CIILotNo { get; set; } = string.Empty;
        
        public double Amount => Math.Round(Price * Qty, 4);
    }

    public class CustomPListModel
    {
        public int ID { get; set; }
        public int OrderItemID { get; set; } // FK to CustomInvoiceItems.EntryID
        public int CartonFrom { get; set; }
        public int CartonTo { get; set; }
        public int CartonQty { get; set; }
        public int Qty { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string BatchNo { get; set; } = string.Empty;
        public long DP_RefID { get; set; } // Tracks Dispatch List EntryID

        // View model helpers
        public string OrderNo { get; set; } = string.Empty;
        public string ProformaInvoice { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    public class CustomInvoiceOtherChargeModel
    {
        public int EntryID { get; set; }
        public int SNo { get; set; }
        public string ChargesDetail { get; set; } = string.Empty;
        public double ChargesAmt { get; set; }
        public bool Custom { get; set; }
    }

    public class AdvancePaymentModel
    {
        public int EntryID { get; set; }
        public string AdviceNo { get; set; } = string.Empty;
        public double Amount { get; set; }
        public double AdjustedAmount { get; set; }
        
        // From FCustAdvancePaymentsAdjusted table where RefID = EntryID
        public double AdjustNow { get; set; }
        public double Balance => Amount - AdjustedAmount - AdjustNow;
    }
}
