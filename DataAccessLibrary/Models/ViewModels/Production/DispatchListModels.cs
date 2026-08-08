using System;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class DispatchListSummaryItem
    {
        public long EntryID { get; set; }
        public string DispatchListNo { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public int TotalCartons { get; set; }
        public decimal TotalQty { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public bool Finalyzed { get; set; }
        public DateTime? FinalyzedDT { get; set; }
        public string FinalyzedUserName { get; set; } = string.Empty;
        public bool NewFormat { get; set; }
        public string UserName { get; set; } = string.Empty;
        public DateTime? EntryDT { get; set; }
    }

    public class DispatchListSearchFilter
    {
        public string CustCode { get; set; } = string.Empty;
        public int DateRangeIndex { get; set; } = 0; // 0=Today, 1=7d, 2=15d, 3=30d, 4=60d, 5=90d, 6=1y, 7=5y, 8=Custom
        public DateTime DtFrom { get; set; } = DateTime.Today;
        public DateTime DtTo { get; set; } = DateTime.Today;
        public bool OnlyNotFinalized { get; set; } = false;
    }

    public class FinalizeDispatchResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
    }

    public class EditDispatchListHeaderDto
    {
        public long EntryID { get; set; }
        public string DispatchListNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public bool Finalyzed { get; set; }
    }
}
