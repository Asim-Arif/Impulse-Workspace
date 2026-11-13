using System;
using System.Collections.Generic;
using System.Linq;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class PpcOrderSummaryListItem
    {
        public string OrderNo { get; set; } = string.Empty;
        public DateTime OrderDate { get; set; }
        public DateTime? DeliveryDate { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public int TotalItems { get; set; }
        public int PlannedItems { get; set; }
        public int TotalOrderQty { get; set; }
        public bool IsPlanned { get; set; }
        public bool IsPartiallyPlanned => !IsPlanned && PlannedItems > 0;
        public string? PlannedBy { get; set; }
        public DateTime? PlannedAt { get; set; }
    }

    public class PpcOrderHeaderDto
    {
        public string OrderNo { get; set; } = string.Empty;
        public DateTime OrderDate { get; set; }
        public DateTime? DeliveryDate { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Currency { get; set; } = string.Empty;
        public bool Authorized { get; set; }
        public bool IsPlanned { get; set; }
        public bool HasDeltaItems => Items.Any(i => i.HasDeltaQty);
        public bool IsPartiallyPlanned => !IsPlanned && (Items.Any(i => i.IsPlanned) || HasDeltaItems);
        public bool HasUnplannedItems => Items.Any(i => (!i.IsPlanned && !i.HasDeltaQty) || !i.Authorized || i.HasDeltaQty);
        public int PlannedItemsCount => Items.Count(i => i.IsPlanned && !i.HasDeltaQty);
        public int ActionRequiredItemsCount => Items.Count(i => i.Authorized && i.NeedsPlanningAction);
        public int LockedPlannedItemsCount => Items.Count(i => i.Authorized && !i.NeedsPlanningAction);
        public int AwaitingAuthItemsCount => Items.Count(i => !i.Authorized);
        public string? PlannedBy { get; set; }
        public DateTime? PlannedAt { get; set; }
        public string? Notes { get; set; }
        public List<PpcOrderItemDto> Items { get; set; } = new();
        public List<PpcPlanningExceptionLogDto> ExceptionLogs { get; set; } = new();
    }

    public class PpcOrderItemDto
    {
        public string ItemCode { get; set; } = string.Empty;
        public string CompItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int OrderQty { get; set; }
        public bool Authorized { get; set; } = true;
        public int AuthorizedQty { get; set; }
        public bool IsSavedPlan { get; set; } = false;
        public int PreviouslyPlannedOrderQty { get; set; } = 0;
        public bool HasDeltaQty => Authorized && OrderQty > PreviouslyPlannedOrderQty && PreviouslyPlannedOrderQty > 0;
        public int DeltaQty => Math.Max(0, OrderQty - PreviouslyPlannedOrderQty);
        private bool _isPlanned;
        public bool IsPlanned
        {
            get => _isPlanned || IsSavedPlan || StockQty > 0 || SFStockQty > 0 || Purchases.Any(p => p.PurchaseQty > 0 && p.ProcessID > 0) || PreviouslyPlannedOrderQty > 0;
            set => _isPlanned = value;
        }
        public int AvailableInHandStock { get; set; }

        // Live execution & exception tracking
        public bool IsUnlockedForRevision { get; set; } = false;
        public bool IsExpandedInLockedTab { get; set; } = false;
        public int AlreadyAdjustedStockQty { get; set; } = 0;
        public int RejectedStockQty { get; set; } = 0;
        public int EffectiveStockQty => Math.Max(0, StockQty - RejectedStockQty);
        public int MinAllowedStockQty => Math.Max(AlreadyAdjustedStockQty, RejectedStockQty);
        public int RejectedSFQty { get; set; } = 0;
        public int AlreadyIssuedSFQty => SFStockAllocations.Sum(a => a.AlreadyIssuedQty);
        public int EffectiveSFStockQty => Math.Max(0, SFStockQty - RejectedSFQty);
        public List<PpcPlanningExceptionLogDto> ItemExceptions { get; set; } = new();

        // 1. Stock Allocation (Finish Stock)
        public int StockQty { get; set; }

        // 2. Semi-Finished (SF) Stock Allocation
        public int SFStockQty => SFStockAllocations.Sum(a => a.AllocatedQty);
        public int AvailableSFStock => AvailableSFLots.Sum(l => l.AvailableQty);
        public List<PpcItemSFLotDto> AvailableSFLots { get; set; } = new();
        public List<PpcItemSFAllocationDto> SFStockAllocations { get; set; } = new();

        // 3. Multiple Purchase Allocations per Item
        public List<PpcItemPurchaseAllocationDto> Purchases { get; set; } = new();

        // 4. Manufacturing Hubs & Schedules
        public int GroupID { get; set; }
        public string GroupName { get; set; } = string.Empty;
        public List<PpcItemHubScheduleDto> HubSchedules { get; set; } = new();

        // Eligible Purchase Processes for this Item (loaded from Processes_Purchase & ItemProcesses)
        public List<PurchaseProcessLookupDto> AvailablePurchaseProcesses { get; set; } = new();

        // Derived Properties
        public int TotalPurchaseQty => Purchases.Where(p => !string.Equals(p.Status, "Cancelled", StringComparison.OrdinalIgnoreCase)).Sum(p => p.PurchaseQty);
        public int ProductionQty => Math.Max(0, OrderQty - EffectiveStockQty - EffectiveSFStockQty - TotalPurchaseQty);
        public int TotalPlannedQty => EffectiveStockQty + EffectiveSFStockQty + TotalPurchaseQty + ProductionQty;
        public bool IsBalanced => TotalPlannedQty == OrderQty;
        public bool NeedsPlanningAction => !IsSavedPlan || HasDeltaQty || IsUnlockedForRevision || !IsBalanced;
        public bool IsLockedInUi => Authorized && IsSavedPlan && !HasDeltaQty && !IsUnlockedForRevision && IsBalanced;
    }

    public class PpcItemSFLotDto
    {
        public int EntryID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public string BatchNo { get; set; } = string.Empty;
        public int Shelf_RefID { get; set; }
        public string StoreLocation { get; set; } = string.Empty;
        public int AvailableQty { get; set; }
    }

    public class PpcItemSFAllocationDto
    {
        public int EntryID { get; set; }
        public int SOO_RefID { get; set; }
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public string BatchNo { get; set; } = string.Empty;
        public int Shelf_RefID { get; set; }
        public string StoreLocation { get; set; } = string.Empty;
        public int AvailableQty { get; set; }
        public int AllocatedQty { get; set; }
        public int AlreadyIssuedQty { get; set; } = 0;
    }

    public class PpcItemPurchaseAllocationDto
    {
        public int EntryID { get; set; }
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public int? VendID { get; set; }
        public string? MakerName { get; set; }
        public int PurchaseQty { get; set; }
        public double PurchaseRate { get; set; }
        public double ExcessQtyPercentage { get; set; } = 50;
        public long? VendIssued_RefID { get; set; }
        public string? MasterPONo { get; set; }
        public DateTime? IssuedAt { get; set; }
        public string? IssuedBy { get; set; }
        public bool IsIssued => VendIssued_RefID.HasValue && VendIssued_RefID.Value > 0;
        public long? VendIssdDetail_EntryID { get; set; }
        public int IssuedQty { get; set; }
        public int ReceivedQty { get; set; }
        public int PendingPoQty => Math.Max(0, IssuedQty - ReceivedQty);
        public string Status { get; set; } = "Active";
        public int? OriginalPlannedQty { get; set; }
        public int ShortClosedQty { get; set; }
        public string? ExceptionNotes { get; set; }
        public bool CanCancelIssuedPo => IsIssued && ReceivedQty <= 0 && !string.Equals(Status, "Cancelled", StringComparison.OrdinalIgnoreCase);
        public bool CanShortCloseIssuedPo => IsIssued && ReceivedQty > 0 && ReceivedQty < PurchaseQty && !string.Equals(Status, "Cancelled", StringComparison.OrdinalIgnoreCase);
        public List<MakerLookupDto> AvailableMakers { get; set; } = new();
    }

    public class PpcItemHubScheduleDto
    {
        public string Hub_Name { get; set; } = string.Empty;
        public DateTime StartDate { get; set; } = DateTime.Today;
        public DateTime EndDate { get; set; } = DateTime.Today.AddDays(7);
        public int PlannedQty { get; set; }
        public List<HubSupervisorDto> Supervisors { get; set; } = new();
    }

    public class PurchaseProcessLookupDto
    {
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
    }

    public class MakerLookupDto
    {
        public int VendID { get; set; }
        public string MakerName { get; set; } = string.Empty;
        public double Rate { get; set; }
    }

    public class PpcPlanSaveRequest
    {
        public string OrderNo { get; set; } = string.Empty;
        public string PlannedBy { get; set; } = string.Empty;
        public string? Notes { get; set; }
        public List<PpcOrderItemDto> Items { get; set; } = new();
    }

    public class PpcPlanningExceptionLogDto
    {
        public int EntryID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public string ExceptionType { get; set; } = string.Empty;
        public int AffectedQty { get; set; }
        public string? SourceRef { get; set; }
        public string Reason { get; set; } = string.Empty;
        public string CreatedBy { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
    }

    public class PpcExceptionActionRequest
    {
        public string OrderNo { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public string ExceptionType { get; set; } = string.Empty;
        public int AffectedQty { get; set; }
        public int? PurchaseSplitEntryID { get; set; }
        public int? SfAllocationSooRefID { get; set; }
        public string Reason { get; set; } = string.Empty;
        public string ActionBy { get; set; } = string.Empty;
    }

    public class PpcExceptionActionResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}
