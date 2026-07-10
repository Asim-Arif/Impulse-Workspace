using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CustomPaymentStatusModel
    {
        public long EntryID { get; set; }
        public int EntryType { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string CustomInvoice { get; set; } = string.Empty;
        public string Orders { get; set; } = string.Empty;
        public string FormE { get; set; } = string.Empty;
        public string FDBCNo { get; set; } = string.Empty;
        public DateTime? FDBCDT { get; set; }
        public DateTime? DT { get; set; }
        public decimal TotalCustomAmt { get; set; }
        public decimal CustomAmtToRcv { get; set; }
        public decimal TotalCommercialAmt { get; set; }
        public decimal TotalCommercialAmtToRcv { get; set; }
        public string Description { get; set; } = string.Empty;
        public string PRCVchrNo { get; set; } = string.Empty;
        public decimal ExchRate { get; set; }
    }
}
