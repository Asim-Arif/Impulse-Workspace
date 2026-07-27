using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class SFTransactionModel
    {
        public int EntryID { get; set; }
        public string UserName { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty; // Process Description
        public int QtyIssd { get; set; }
        public int Issuance { get; set; } // 0 = Receiving, 1 = Issuance
        public string Location { get; set; } = string.Empty;
        public string StoreName { get; set; } = string.Empty;
        public int Store_RefID { get; set; }
        public string LotNo { get; set; } = string.Empty;

        public int SignedQty => Issuance == 0 ? QtyIssd : -QtyIssd;
        public string TypeName => Issuance == 0 ? "Receiving" : "Issuance";
    }

    public class SFTransactionFilterModel
    {
        public DateTime FromDate { get; set; } = DateTime.Today;
        public DateTime ToDate { get; set; } = DateTime.Today;
        public bool IncludeIssuance { get; set; } = true;
        public bool IncludeReceiving { get; set; } = true;
        public int StoreRefId { get; set; } = 0;
        public string ItemCode { get; set; } = "0";
        public int ProcessId { get; set; } = 0;
    }

    public class ProcessLookupModel
    {
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
        public string Code { get; set; } = string.Empty;
        public string DisplayText => string.IsNullOrWhiteSpace(Code) ? Description : $"{Description} ({Code})";
    }
}
