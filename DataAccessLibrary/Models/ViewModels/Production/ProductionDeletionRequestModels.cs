using System;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class ProductionDeletionRequestModel
    {
        public int Id { get; set; }
        public string RequestType { get; set; } = "LotReceiving"; // 'LotReceiving', 'LotIssuance', 'MasterPOIssuance', 'SkipProcess'
        public long EntityRefID { get; set; }                     // VRD_EntryID, Issuance EntryID, etc.
        public string LotNo { get; set; } = string.Empty;
        public string? OrderNo { get; set; }
        public string? ItemCode { get; set; }
        public string? ItemName { get; set; }
        public int? ProcessID { get; set; }
        public string? ProcessName { get; set; }
        public long? MakerID { get; set; }
        public string? MakerName { get; set; }
        public decimal Qty { get; set; }
        public string RequestedBy { get; set; } = string.Empty;
        public DateTime RequestedDT { get; set; } = DateTime.Now;
        public string Reason { get; set; } = string.Empty;
        public string? MachineName { get; set; }
        public string Status { get; set; } = "Pending"; // 'Pending', 'Approved', 'Rejected', 'Cancelled'
        public string? ReviewedBy { get; set; }
        public DateTime? ReviewedDT { get; set; }
        public string? DirectorRemarks { get; set; }
        public int? TaskId { get; set; }

        public int TargetProcessID
        {
            get => (int)(MakerID ?? 0);
            set => MakerID = value;
        }

        public string TargetProcessName
        {
            get => MakerName ?? string.Empty;
            set => MakerName = value;
        }
    }
}
