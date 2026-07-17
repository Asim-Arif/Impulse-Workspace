using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class TotalExportModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string CustomInvoice { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public double TotalAmt { get; set; }
        public double TotalCustomAmt { get; set; }
        public double ExchRate { get; set; }
        public string GatePassNo { get; set; } = string.Empty;
    }
}
