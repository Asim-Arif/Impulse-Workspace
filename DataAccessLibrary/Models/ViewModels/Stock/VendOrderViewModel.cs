using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendOrderViewModel
    {
        public string OrderNo { get; set; }
        public DateTime? DT { get; set; } = DateTime.Now;
        public string VendID { get; set; }
        public string AttnPerson { get; set; }
        public string PaymentTerms { get; set; }
        public string Department { get; set; }
        public DateTime? DeliveryDT { get; set; } = DateTime.Now;
        public string Remarks { get; set; }
        public string ReqNo { get; set; }
        public bool? VeryUrgent { get; set; } = false;
        public bool? Final { get; set; }
        public int? RunningPONo { get; set; }
        public int EntryID { get; set; }
        public byte? IssuanceType { get; set; }
        public string Ref { get; set; }
        public string Payments { get; set; }
        public string Shipment { get; set; }
        public string Instructions { get; set; }
        public string Notes { get; set; }
        public string Currency { get; set; }
        public bool? SampleProvided { get; set; } = false;
        public string SampleProvidedDescription { get; set; }
        public bool? DrawingProvided { get; set; } = false;
        public string DrawingProvidedDescription { get; set; }
        public bool? SteelProvided { get; set; } = false;
        public string SteelProvidedDescription { get; set; }
        public byte? POType { get; set; }
        public string CustomerPONo { get; set; }
        public int? SteelType_RefID { get; set; }
        public bool? SampleOrder { get; set; } = false;
        public string Order_Collection_By { get; set; }

        public List<VendOrderDetailViewModel> OrderDetails { get; set; } = new List<VendOrderDetailViewModel>();
    }
}
