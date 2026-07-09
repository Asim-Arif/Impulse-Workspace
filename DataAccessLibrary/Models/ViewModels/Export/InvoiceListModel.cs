using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class InvoiceListModel
    {
        public string CustomInvoice { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string Curr { get; set; } = string.Empty;
        public double TotalAmt { get; set; }
        public double TotalCustomAmt { get; set; }
        public int PackingInfo { get; set; }
        public string ShippingMode { get; set; } = string.Empty;
        public DateTime? GatePassDT { get; set; }
        public DateTime? DTMaturity { get; set; }
        public bool SampleInvoice { get; set; }
        public string InvoiceOrders { get; set; } = string.Empty;
        public string InvoiceInternalOrders { get; set; } = string.Empty;
        public string DispatchListNos { get; set; } = string.Empty;
    }
}
