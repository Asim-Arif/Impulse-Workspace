using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class OrderLotsTrackingFilter
    {
        public string? CustCode { get; set; }
        public string? OrderNo { get; set; }
        public string? InternalRefNo { get; set; }
        public string? HubName { get; set; }
        public int DateRangeType { get; set; } = 3; // 0=All, 1=Today, 2=Last 15d, 3=Last 30d, 4=Last 60d, 5=Last 90d, 6=Custom
        public int DateFilterMode { get; set; } = 0; // 0=Order Date, 1=Target Date
        public DateTime? DtFrom { get; set; } = DateTime.Today.AddDays(-30);
        public DateTime? DtTo { get; set; } = DateTime.Today;
        public bool IncludeCompleted { get; set; } = false; // Default: Active WIP lots only
        public string? SearchText { get; set; }
    }

    public class OrderLotTrackingItemDto
    {
        public string LotNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string ItemSize { get; set; } = string.Empty;
        public string FullArticle => $"{ItemCode} {ItemName} {ItemSize}".Trim();

        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string DisplayCustomer => CustCode;

        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public string Hub_Name { get; set; } = string.Empty;
        public string Supervisors { get; set; } = string.Empty;
        public int Qty { get; set; }

        public int CurrentSeqNo { get; set; } = 1;
        public int TotalSeqNo { get; set; } = 1;
        public double ProgressPct => TotalSeqNo > 0 ? Math.Min(100.0, Math.Round((CurrentSeqNo * 100.0) / TotalSeqNo, 1)) : 0;

        public DateTime? TargetDate { get; set; }
        public DateTime? PlannedStartDate { get; set; }
        public DateTime? OrderDeliveryDate { get; set; }
        public DateTime? LastActivityDT { get; set; }
        public string MakerName { get; set; } = string.Empty;

        public string LotState { get; set; } = string.Empty; // "Issued to Maker", "In Hub (Received)", "Completed"
        public bool IsCompleted { get; set; }
        public bool IsReWork { get; set; }

        public int? DaysRemaining => TargetDate.HasValue ? (TargetDate.Value.Date - DateTime.Today).Days : null;

        public string TargetStatus
        {
            get
            {
                if (IsCompleted) return "Completed";
                if (!TargetDate.HasValue) return "No Target";
                int days = (TargetDate.Value.Date - DateTime.Today).Days;
                if (days < 0) return "Overdue";
                if (days == 0) return "Due Today";
                if (days <= 3) return "Due Soon";
                return "On Track";
            }
        }

        public string TargetStatusBadgeClass => TargetStatus switch
        {
            "Completed" => "bg-success text-white",
            "Overdue" => "bg-danger text-white",
            "Due Today" => "bg-warning text-dark",
            "Due Soon" => "bg-warning text-dark",
            "On Track" => "bg-info text-dark",
            _ => "bg-secondary text-white"
        };
    }

    public class OrderLotsHubSummaryDto
    {
        public string Hub_Name { get; set; } = string.Empty;
        public int LotCount { get; set; }
        public int TotalQty { get; set; }
        public int OverdueCount { get; set; }
        public string Supervisors { get; set; } = string.Empty;
    }

    public class OrderLotsTrackingDashboardDto
    {
        public int TotalLots { get; set; }
        public int TotalQty { get; set; }
        public int ActiveWipLots { get; set; }
        public int ActiveWipQty { get; set; }
        public int OnTrackLots { get; set; }
        public int DueSoonLots { get; set; }
        public int OverdueLots { get; set; }
        public int CompletedLots { get; set; }

        public List<OrderLotsHubSummaryDto> HubSummaries { get; set; } = new();
        public List<OrderLotTrackingItemDto> Lots { get; set; } = new();
    }

    public class TrackingCustomerLookupItem
    {
        public string CustCode { get; set; } = string.Empty;
        public string CustName { get; set; } = string.Empty;
        public string DisplayCust => CustCode;
    }

    public class TrackingOrderLookupItem
    {
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string DisplayText => !string.IsNullOrEmpty(InternalRefNo) ? InternalRefNo : OrderNo;
    }
}
