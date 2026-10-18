using System;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class WorkflowApprovalHistoryModel
    {
        public string SourceTable { get; set; } = string.Empty;
        public int RecordId { get; set; }
        public int? TaskId { get; set; }
        public string WorkflowCode { get; set; } = string.Empty;
        public string WorkflowName { get; set; } = string.Empty;
        public string Module { get; set; } = string.Empty;
        public string? LotNo { get; set; }
        public string? OrderNo { get; set; }
        public string? ItemCode { get; set; }
        public string? ItemName { get; set; }
        public string? ProcessName { get; set; }
        public string? MakerName { get; set; }
        public string? TargetProcessName { get; set; }
        public decimal? Qty { get; set; }
        public string? RequestReason { get; set; }
        public string? RequestedBy { get; set; }
        public DateTime? RequestedDT { get; set; }
        public string? MachineName { get; set; }
        public string Status { get; set; } = "Pending";
        public string? ActionBy { get; set; }
        public DateTime? ActionDT { get; set; }
        public string? ActionRemarks { get; set; }

        public string DisplayIdentifier
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(LotNo) && LotNo != "0")
                {
                    return $"Lot #{LotNo}" + (!string.IsNullOrWhiteSpace(OrderNo) ? $" ({OrderNo})" : "");
                }
                if (!string.IsNullOrWhiteSpace(OrderNo))
                {
                    return $"PO #{OrderNo}";
                }
                if (!string.IsNullOrWhiteSpace(ItemName))
                {
                    return ItemName;
                }
                return $"Ref #{RecordId}";
            }
        }
    }
}
