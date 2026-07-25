using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class RMPOListViewModel
    {
        public DateTime? DT { get; set; }
        public string VendID { get; set; }
        public string AccTitle { get; set; }
        public string AttnPerson { get; set; }
        public string PaymentTerms { get; set; }
        public string Department { get; set; }
        public DateTime? DeliveryDT { get; set; }
        public string Remarks { get; set; }
        public string OrderNo { get; set; }
        public string ReqNo { get; set; }
        public bool? Final { get; set; }
        public int? RunningPONo { get; set; }
        public byte? IssuanceType { get; set; }
        public byte[] PDFAttachment { get; set; }
        public double? TotalQtyOrdered { get; set; }
        public bool? SampleOrder { get; set; }
        public string RcvableOrderNo { get; set; }
        public double? TotalQtyRcvd { get; set; }
    }
}
