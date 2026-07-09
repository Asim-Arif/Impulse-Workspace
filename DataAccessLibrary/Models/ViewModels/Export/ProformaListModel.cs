using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class ProformaListModel
    {
        public string PInvoice { get; set; } = string.Empty;
        public string CompanyName { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string Curr { get; set; } = string.Empty;
        public double ProformaAmt { get; set; }
        public double AdvanceAmount { get; set; }
        public double ChargesAmt { get; set; }
        public double DiscountAmount { get; set; }
        public double Insurance_Freight_Amount { get; set; }
        public double Insurance_Freight_Amount_II { get; set; }
        public string Orders { get; set; } = string.Empty;
        public string InternalOrders { get; set; } = string.Empty;
        public string AdvanceDescription { get; set; } = string.Empty;
        public long CompanyRefID { get; set; }

        // Calculated fields for UI
        public double CalculatedProformaAmt => ProformaAmt + ChargesAmt - DiscountAmount + Insurance_Freight_Amount + Insurance_Freight_Amount_II;
    }
}
