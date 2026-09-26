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
        public int TotalOrderQty { get; set; }
        public bool IsPlanned { get; set; }
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
        public string? PlannedBy { get; set; }
        public DateTime? PlannedAt { get; set; }
        public string? Notes { get; set; }
        public List<PpcOrderItemDto> Items { get; set; } = new();
    }

    public class PpcOrderItemDto
    {
        public string ItemCode { get; set; } = string.Empty;
        public string CompItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int OrderQty { get; set; }
        public int AvailableInHandStock { get; set; }

        // 1. Stock Allocation
        public int StockQty { get; set; }

        // 2. Multiple Purchase Allocations per Item
        public List<PpcItemPurchaseAllocationDto> Purchases { get; set; } = new();

        // 3. Manufacturing Hubs & Schedules
        public int GroupID { get; set; }
        public string GroupName { get; set; } = string.Empty;
        public List<PpcItemHubScheduleDto> HubSchedules { get; set; } = new();

        // Eligible Purchase Processes for this Item (loaded from Processes_Purchase & ItemProcesses)
        public List<PurchaseProcessLookupDto> AvailablePurchaseProcesses { get; set; } = new();

        // Derived Properties
        public int TotalPurchaseQty => Purchases.Sum(p => p.PurchaseQty);
        public int ProductionQty => Math.Max(0, OrderQty - StockQty - TotalPurchaseQty);
        public int TotalPlannedQty => StockQty + TotalPurchaseQty + ProductionQty;
        public bool IsBalanced => (StockQty + TotalPurchaseQty + ProductionQty) == OrderQty;
    }

    public class PpcItemPurchaseAllocationDto
    {
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public int? VendID { get; set; }
        public string? MakerName { get; set; }
        public int PurchaseQty { get; set; }
        public double PurchaseRate { get; set; }
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
}
