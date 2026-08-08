using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Dashboard
{
    public class ProductionPlanningFilters
    {
        public string CustCode { get; set; } = "0";
        public string Country { get; set; } = "<Comp>";
        public string ItemCode { get; set; } = string.Empty;
        public string? FilterOrderNo { get; set; }
    }

    public class PlanningEquationSummary
    {
        public long TotalCustOrders { get; set; }
        public long MinStock { get; set; }
        public long TotalRunningLots { get; set; }
        public long TotalOpenPO { get; set; }
        public long TotalSFStock { get; set; }
        public long TotalRFStock { get; set; }
        public long TotalForging { get; set; }
        public long QtyRequired { get; set; }
        public double CompletionPercentage { get; set; }

        public bool HasDeficit => QtyRequired > 0;
        public long EffectiveDemand => TotalCustOrders + MinStock;
        public long TotalSuppliesInPipeline => TotalRunningLots + TotalOpenPO + TotalSFStock + TotalRFStock;
    }

    public class CustOrderDemandItem
    {
        public long ID { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public long Qty { get; set; }
        public long ShippedQty { get; set; }
        public long UnshippedQty => Math.Max(0, Qty - ShippedQty);
        public DateTime? DeliveryDT { get; set; }
        public int DeliveryStatus { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public bool IsUrgent => DeliveryStatus == 2 || Remarks.IndexOf("Urgent", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    public class RunningLotItem
    {
        public string LotNo { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public long Qty { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public DateTime? R2InDT { get; set; }
        public string OriginType { get; set; } = string.Empty; // "Semi", "Pack", or ""
        public long IssuanceEntryID { get; set; }
        public long RcvEntryID { get; set; }
        public long ProcessID { get; set; }
    }

    public class OpenPOItem
    {
        public long EntryID { get; set; }
        public string ReceiptID { get; set; } = string.Empty;
        public string MakerName { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public long Qty { get; set; }
        public DateTime? ReturnDT { get; set; }
        public string MasterPONo { get; set; } = string.Empty;
        public bool IsOverdue => ReturnDT.HasValue && ReturnDT.Value.Date < DateTime.Today;
    }

    public class SemiFinishStockItem
    {
        public string Description { get; set; } = string.Empty;
        public long Qty { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public string LotNoManual { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public bool IsLookAlike { get; set; }
    }

    public class ReadyFinishStockItem
    {
        public string ItemID { get; set; } = string.Empty;
        public long NetQty { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public bool IsLookAlike { get; set; }
    }

    public class ForgingRawMaterialItem
    {
        public long RMID { get; set; }
        public string MaterialID { get; set; } = string.Empty;
        public long ShelfQty { get; set; }
        public string LocationDisplay { get; set; } = string.Empty;
        public bool IsInactive { get; set; }
    }

    public class ForgingCrossDemandItem
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public long UnshippedQty { get; set; }
    }

    public class ItemProcessRoutingItem
    {
        public int SNo { get; set; }
        public string Description { get; set; } = string.Empty;
        public decimal Rate { get; set; }
    }

    public class MakerRateAssignmentItem
    {
        public string MakerName { get; set; } = string.Empty;
        public string ProcessDescription { get; set; } = string.Empty;
        public string RateDisplay { get; set; } = string.Empty;
    }

    public class CustomerCatalogRefItem
    {
        public string CustCode { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public string OCR { get; set; } = string.Empty;
        public long OrderedQty { get; set; }
    }

    public class CustomerComplaintItem
    {
        public long EntryID { get; set; }
        public string ComplaintNo { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    public class ProductionPlanningDashboardData
    {
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public PlanningEquationSummary Summary { get; set; } = new();

        public List<CustOrderDemandItem> CustomerOrders { get; set; } = new();
        public List<RunningLotItem> RunningLots { get; set; } = new();
        public List<OpenPOItem> OpenPOs { get; set; } = new();
        public List<SemiFinishStockItem> SemiFinishStocks { get; set; } = new();
        public List<SemiFinishStockItem> SemiFinishLookalikes { get; set; } = new();
        public List<ReadyFinishStockItem> ReadyFinishStocks { get; set; } = new();
        public List<ReadyFinishStockItem> ReadyFinishLookalikes { get; set; } = new();
        public List<ForgingRawMaterialItem> Forgings { get; set; } = new();
        public List<ItemProcessRoutingItem> ProcessRoutings { get; set; } = new();
        public List<MakerRateAssignmentItem> MakerRates { get; set; } = new();
        public List<CustomerCatalogRefItem> CustomerCatalogs { get; set; } = new();
        public List<CustomerComplaintItem> Complaints { get; set; } = new();
    }
}
