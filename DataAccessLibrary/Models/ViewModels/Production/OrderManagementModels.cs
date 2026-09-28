using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class OrderManagementFilter
    {
        public string CustCode { get; set; } = string.Empty;
        public int DateRangeType { get; set; } = 0; // 0=All, 1=Last 30 Days, 2=Last 90 Days, 3=This Year, 4=Custom
        public DateTime? FromDate { get; set; }
        public DateTime? ToDate { get; set; }
        public string SearchText { get; set; } = string.Empty;
        public string StatusFilter { get; set; } = "All"; // All, InProgress, Completed, Overdue
    }

    public class OrderSummaryCardDto
    {
        public int TotalOrders { get; set; }
        public int InProgressOrders { get; set; }
        public int CompletedOrders { get; set; }
        public int OverdueOrders { get; set; }
        public int TotalOrderedPcs { get; set; }
        public int TotalProducedPcs { get; set; }
        public int TotalDispatchedPcs { get; set; }
        public double AverageProgressPct { get; set; }
    }

    public class CustomerOrderHeaderDto
    {
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public DateTime? DeliveryDT { get; set; }
        public int TotalOrderQty { get; set; }
        public int TotalPlannedQty { get; set; }
        public int TotalShippedQty { get; set; }
        public int TotalArticles { get; set; }
        public bool Authorized { get; set; }
        public bool OrderPlanApproved { get; set; }
        public double ProgressPct { get; set; }
        public int CurrentStageNumber { get; set; }
        public string CurrentStageName { get; set; } = string.Empty;

        public bool IsOverdue => DeliveryDT.HasValue && DeliveryDT.Value.Date < DateTime.Today && ProgressPct < 100;

        public string StatusBadgeCss => ProgressPct >= 100 ? "badge-completed"
            : IsOverdue ? "badge-overdue"
            : CurrentStageNumber >= 8 ? "badge-staging"
            : CurrentStageNumber >= 4 ? "badge-inprogress"
            : "badge-planned";

        public string StatusText => ProgressPct >= 100 ? "Completed"
            : IsOverdue ? "Delayed"
            : CurrentStageNumber >= 8 ? "Dispatch Ready"
            : CurrentStageNumber >= 4 ? "In Production"
            : CurrentStageNumber >= 3 ? "PPC Planned"
            : CurrentStageNumber >= 2 ? "Authorized"
            : "Order Entered";
    }

    public class OrderItemProgressDto
    {
        public int ID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string CompItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int OrderedQty { get; set; }
        public int PlannedQty { get; set; }
        public int ProducedQty { get; set; }
        public int DispatchedQty { get; set; }
        public int RemainingQty => Math.Max(0, OrderedQty - ProducedQty);
        public DateTime? DeliveryDT { get; set; }
        public string Packaging { get; set; } = string.Empty;
        public string Quality { get; set; } = string.Empty;
        public int GroupID { get; set; }
        public string BatchNo { get; set; } = string.Empty;

        // PPC Planning & Stock Adjustment Quantities
        public int PlannedStockQty { get; set; }
        public int PlannedPurchaseQty { get; set; }
        public int PlannedProductionQty { get; set; }
        public int StockAdjustedQty { get; set; }
        public bool HasPpcPlan => PlannedPurchaseQty > 0 || PlannedStockQty > 0 || PlannedProductionQty > 0 || StockAdjustedQty > 0;

        // 10-Stage Workflow Metrics
        public int CurrentStageNumber { get; set; } = 1;
        public string CurrentStageName { get; set; } = "Order Entered";
        public double ProgressPct { get; set; }

        public bool IsOverdue => DeliveryDT.HasValue && DeliveryDT.Value.Date < DateTime.Today && ProgressPct < 100;

        public string PackagingBadgeCss => Packaging?.ToLower() switch
        {
            "completed" => "badge-completed",
            "inprogress" => "badge-inprogress",
            "staged" => "badge-staging",
            _ => "badge-muted"
        };

        public string StageBadgeCss => CurrentStageNumber switch
        {
            >= 10 => "badge-completed",
            >= 8 => "badge-staging",
            >= 6 => "badge-inprogress",
            >= 4 => "badge-production",
            _ => "badge-planned"
        };
    }

    public class ItemPurchaseOrderDto
    {
        public int EntryID { get; set; }
        public string POReceiptID { get; set; } = string.Empty;
        public string MasterPONo { get; set; } = string.Empty;
        public int VendID { get; set; }
        public string VenderName { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public int IssQty { get; set; }
        public int RcvdQty { get; set; }
        public int Balance => Math.Max(0, IssQty - RcvdQty);
        public double ProgressPct => IssQty > 0 ? Math.Min(100.0, Math.Round((RcvdQty * 100.0) / IssQty, 1)) : 0;
        public DateTime? ReturnDT { get; set; }
        public bool IsDelayed => ReturnDT.HasValue && ReturnDT.Value.Date < DateTime.Today && Balance > 0;
        public string StatusText => Balance == 0 && IssQty > 0 ? "Fulfilled" : RcvdQty > 0 ? "Partial" : "Issued";
    }

    public class ItemRunningLotDto
    {
        public string LotNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public int Qty { get; set; }
        public int CurrentSeqNo { get; set; }
        public int TotalSeqNo { get; set; }
        public double ProgressPct => TotalSeqNo > 0 ? Math.Min(100.0, Math.Round((CurrentSeqNo * 100.0) / TotalSeqNo, 1)) : 0;
        public DateTime? LastActivityDT { get; set; }
        public string MakerName { get; set; } = string.Empty;
        public bool ReWorkLot { get; set; }
    }

    public class ItemDispatchDetailDto
    {
        public int DispatchID { get; set; }
        public string DispatchListNo { get; set; } = string.Empty;
        public int CartonNo { get; set; }
        public int InnerNo { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public int Qty { get; set; }
        public DateTime? EntryDT { get; set; }
        public string AddedBy { get; set; } = string.Empty;
        public bool Finalyzed { get; set; }
        public DateTime? FinalyzedDT { get; set; }
        public string InvoiceNo { get; set; } = string.Empty;
        public DateTime? InvoiceDate { get; set; }
        public string StatusText => !string.IsNullOrEmpty(InvoiceNo) ? "Invoiced" : Finalyzed ? "Finalized" : "Staged (Draft)";
    }

    public class ItemStockAdjustmentDto
    {
        public int EntryID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public int Qty { get; set; }
        public string UserName { get; set; } = string.Empty;
        public DateTime? DTEntry { get; set; }
        public int Shelf_RefID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public int? NextProcessID { get; set; }
        public string StartingProcessName { get; set; } = string.Empty;
    }
}
