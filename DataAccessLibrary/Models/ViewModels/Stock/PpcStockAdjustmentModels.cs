using System;
using System.Collections.Generic;
using System.Linq;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class PpcStockAdjustmentOrderHeaderModel
    {
        public string OrderNo { get; set; } = string.Empty;
        public DateTime? OrderDate { get; set; }
        public string CustomerCode { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string CustPONo { get; set; } = string.Empty;
        public int TotalPlannedItems { get; set; }
        public int TotalPlannedStockQty { get; set; }
        public int TotalRemainingStockQty { get; set; }
    }

    public class ItemProcessLookupModel
    {
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
        public decimal ItemSNo { get; set; }
    }

    public class PpcStockShelfAllocationModel
    {
        public int Shelf_RefID { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public string Batch_No { get; set; } = string.Empty;
        public string Mill_Certificate_No { get; set; } = string.Empty;
        public int RCV_ISD_RefID { get; set; }
        public double NetAvailableQty { get; set; }
        public int AllocatedQty { get; set; }

        public string LocationDisplayText =>
            $"{StoreName} {RackNo} {ShelfNo} (Avail: {NetAvailableQty}) [Lot: {LotNo}] [Batch: {Batch_No}]";
    }

    public class PpcStockAdjustmentPlannedItemModel
    {
        public int EntryID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int OrderQty { get; set; }
        public int PlannedStockQty { get; set; }
        public int AlreadyAdjustedQty { get; set; }
        public int RemainingQtyToIssue => Math.Max(0, PlannedStockQty - AlreadyAdjustedQty);
        public double AvailableInHandQty { get; set; }

        public int AdjQty { get; set; }
        public int? SelectedProcessId { get; set; }

        public List<ItemProcessLookupModel> Processes { get; set; } = new();
        public List<PpcStockShelfAllocationModel> AvailableShelfLots { get; set; } = new();
        public List<PpcStockShelfAllocationModel> Allocations { get; set; } = new();

        public int TotalAllocatedQty => Allocations?.Sum(a => a.AllocatedQty) ?? 0;
        public bool IsAllocationValid => AdjQty > 0 && TotalAllocatedQty == AdjQty && SelectedProcessId.HasValue && SelectedProcessId.Value > 0;
    }

    public class PpcStockAdjustmentSubmissionModel
    {
        public string OrderNo { get; set; } = string.Empty;
        public DateTime AdjustmentDate { get; set; } = DateTime.Today;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
        public List<PpcStockAdjustmentPlannedItemModel> Items { get; set; } = new();
    }

    public class PpcStockAdjustmentResultModel
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public List<string> GeneratedLots { get; set; } = new();
    }
}
