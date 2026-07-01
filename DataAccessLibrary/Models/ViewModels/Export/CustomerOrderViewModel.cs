using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CustomerOrderViewModel
    {
        public string OrderNo { get; set; } = string.Empty;
        public DateTime DT { get; set; } = DateTime.Today;
        public string TradeTerms { get; set; } = "FOB-Sialkot";
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Currency { get; set; } = string.Empty;
        public string PartialShipment { get; set; } = "Allowed";
        public string PaymentTerms { get; set; } = string.Empty;
        public string TransShipment { get; set; } = "Allowed";
        public string Packaging { get; set; } = string.Empty;
        public DateTime DeliveryDT { get; set; } = DateTime.Today.AddMonths(2);
        public int CompanyRefID { get; set; } = 1;
        public string StampDT { get; set; } = string.Empty;
        public string Quality { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string OrderRcvdVia { get; set; } = string.Empty;
        public string OrderType { get; set; } = "Customer Order";
        public string ShippingMode { get; set; } = string.Empty;
        public int OrderRevisionNo { get; set; } = 0;
        public DateTime OrderRevisionDT { get; set; } = DateTime.Today;
        public string Packaging_Weight { get; set; } = string.Empty;

        // Child collection for grid rows
        public List<CustomerOrderItemViewModel> OrderItems { get; set; } = new();
    }
}
