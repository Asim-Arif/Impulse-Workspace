using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.WorkflowTasks;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Production
{
    public class PpcOrderPlanningService : IPpcOrderPlanningService
    {
        private readonly IPpcOrderPlanningDataAccess _dataAccess;
        private readonly IWorkflowTaskEngine _workflowTaskEngine;
        private readonly ILogger<PpcOrderPlanningService> _logger;

        public PpcOrderPlanningService(
            IPpcOrderPlanningDataAccess dataAccess,
            IWorkflowTaskEngine workflowTaskEngine,
            ILogger<PpcOrderPlanningService> logger)
        {
            _dataAccess = dataAccess;
            _workflowTaskEngine = workflowTaskEngine;
            _logger = logger;
        }

        public async Task<List<PpcOrderSummaryListItem>> GetAuthorizedOrdersForPlanningAsync()
        {
            return await _dataAccess.GetAuthorizedOrdersForPlanningAsync();
        }

        public async Task<PpcOrderHeaderDto?> GetOrderForPlanningAsync(string orderNo)
        {
            return await _dataAccess.GetOrderForPlanningAsync(orderNo);
        }

        public async Task<List<PurchaseProcessLookupDto>> GetPurchaseProcessesForItemAsync(string itemId)
        {
            return await _dataAccess.GetPurchaseProcessesForItemAsync(itemId);
        }

        public async Task<List<MakerLookupDto>> GetMakersForItemAndProcessAsync(string itemId, int processId)
        {
            return await _dataAccess.GetMakersForItemAndProcessAsync(itemId, processId);
        }

        public async Task<bool> SavePpcPlanAndDispatchTasksAsync(PpcPlanSaveRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.OrderNo)) return false;

            bool ok = await _dataAccess.SavePpcPlanAsync(request);
            if (!ok) return false;

            // Complete pending PPC workflow task for this customer order
            try
            {
                await _workflowTaskEngine.CompleteRoleTaskAsync("CustomerOrder", request.OrderNo, "PPC", request.PlannedBy ?? "PPC Planner");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error completing PPC task for Order #{OrderNo}", request.OrderNo);
            }

            // Dispatch 3 streams of tasks and real-time notifications
            try
            {
                await DispatchDownstreamWorkflowTasksAsync(request);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error dispatching downstream workflow tasks for Order #{OrderNo}", request.OrderNo);
            }

            return true;
        }

        public async Task<PpcExceptionActionResult> ExecutePlanningExceptionAsync(PpcExceptionActionRequest request)
        {
            return await _dataAccess.ExecutePlanningExceptionAsync(request);
        }

        private static bool IsPendingUnissuedPurchase(PpcItemPurchaseAllocationDto p) =>
            !p.IsIssued &&
            !string.Equals(p.Status, "Cancelled", StringComparison.OrdinalIgnoreCase) &&
            p.PurchaseQty > 0 &&
            p.ProcessID > 0;

        private async Task DispatchDownstreamWorkflowTasksAsync(PpcPlanSaveRequest request)
        {
            var orderNo = request.OrderNo;
            var plannerUser = request.PlannedBy ?? "PPC Planner";

            // ── 1. Stock User Role Task & Notification (Pending Finish Stock Only) ──
            int totalPendingStockQty = request.Items.Sum(i => Math.Max(0, i.StockQty - i.AlreadyAdjustedStockQty));
            int pendingStockItemsCount = request.Items.Count(i => (i.StockQty - i.AlreadyAdjustedStockQty) > 0);
            bool hasActiveUnissuedStock = request.Items.Any(i => !i.IsLockedInUi && (i.StockQty - i.AlreadyAdjustedStockQty) > 0);

            if (totalPendingStockQty > 0 && hasActiveUnissuedStock)
            {
                try
                {
                    await _workflowTaskEngine.CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                    {
                        SourceEntityType = "CustomerOrder",
                        SourceEntityRefId = orderNo,
                        TargetRole = "Stock",
                        Title = $"Finish Stock Issuance Request for Order #{orderNo}",
                        Description = $"PPC has planned Finish Stock issuance of {totalPendingStockQty:N0} pending pcs across {pendingStockItemsCount} items for Customer Order #{orderNo}. Please review and issue from store.",
                        ActionUrl = $"/stock/stock-order-adjustment?orderNo={orderNo}",
                        Priority = 2, // High
                        DueDate = DateTime.Today.AddDays(2),
                        CreatedBy = plannerUser,
                        UpdateExistingIfExists = true
                    });

                    _logger.LogInformation("Finish stock issuance task dispatched for Order #{OrderNo} ({PendingStockQty} pending pcs)", orderNo, totalPendingStockQty);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to create Stock role task for Order #{OrderNo}", orderNo);
                }
            }

            // ── 1b. Semi-Finished (SF) Stock Role Task & Notification (Pending SF Stock Only) ──
            int totalPendingSFQty = request.Items.Sum(i => i.SFStockAllocations.Sum(a => Math.Max(0, a.AllocatedQty - a.AlreadyIssuedQty)));
            int pendingSfItemsCount = request.Items.Count(i => i.SFStockAllocations.Sum(a => Math.Max(0, a.AllocatedQty - a.AlreadyIssuedQty)) > 0);
            bool hasActiveUnissuedSF = request.Items.Any(i => !i.IsLockedInUi && i.SFStockAllocations.Sum(a => Math.Max(0, a.AllocatedQty - a.AlreadyIssuedQty)) > 0);

            if (totalPendingSFQty > 0 && hasActiveUnissuedSF)
            {
                try
                {
                    await _workflowTaskEngine.CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                    {
                        SourceEntityType = "CustomerOrder",
                        SourceEntityRefId = orderNo,
                        TargetRole = "Stock",
                        Title = $"Semi-Finished Stock Issuance for Order #{orderNo}",
                        Description = $"PPC has planned Semi-Finished stock issuance of {totalPendingSFQty:N0} pending pcs across {pendingSfItemsCount} items for Customer Order #{orderNo}. Please issue via 'Maker Issuance from SF' to return to production.",
                        ActionUrl = $"/production/maker-issuance-from-sf?orderNo={orderNo}&returnUrl={Uri.EscapeDataString($"/production/order-planning?orderNo={orderNo}")}",
                        Priority = 2, // High
                        DueDate = DateTime.Today.AddDays(2),
                        CreatedBy = plannerUser,
                        UpdateExistingIfExists = true
                    });

                    _logger.LogInformation("Semi-finished stock task dispatched for Order #{OrderNo} ({PendingSFQty} pending pcs)", orderNo, totalPendingSFQty);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to create SF Stock role task for Order #{OrderNo}", orderNo);
                }
            }

            // ── 2. Purchaser User Role Task & Notification (Pending Unissued POs Only) ──
            int totalPendingPurchaseQty = request.Items.Sum(i =>
                i.Purchases.Where(IsPendingUnissuedPurchase).Sum(p => p.PurchaseQty));
            int pendingPurchaseItemsCount = request.Items.Count(i =>
                i.Purchases.Any(IsPendingUnissuedPurchase));
            bool hasActiveUnissuedPurchase = request.Items.Any(i =>
                !i.IsLockedInUi && i.Purchases.Any(IsPendingUnissuedPurchase));

            if (totalPendingPurchaseQty > 0 && hasActiveUnissuedPurchase)
            {
                try
                {
                    var purchaseSummaryList = new List<string>();
                    foreach (var item in request.Items.Where(i => i.Purchases.Any(IsPendingUnissuedPurchase)))
                    {
                        foreach (var p in item.Purchases.Where(IsPendingUnissuedPurchase))
                        {
                            string makerPart = !string.IsNullOrEmpty(p.MakerName) ? $" with {p.MakerName}" : "";
                            purchaseSummaryList.Add($"• {item.CompItemID}: {p.PurchaseQty:N0} pcs at [{p.ProcessName}]{makerPart}");
                        }
                    }

                    string summaryText = string.Join("\n", purchaseSummaryList.Take(5));
                    if (purchaseSummaryList.Count > 5) summaryText += $"\n...and {purchaseSummaryList.Count - 5} more";

                    await _workflowTaskEngine.CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                    {
                        SourceEntityType = "CustomerOrder",
                        SourceEntityRefId = orderNo,
                        TargetRole = "Purchaser",
                        Title = $"Maker Purchase Orders Required for Order #{orderNo}",
                        Description = $"PPC has planned Maker POs for {totalPendingPurchaseQty:N0} pending pcs across {pendingPurchaseItemsCount} items for Customer Order #{orderNo}.\n{summaryText}\nPlease generate Maker Purchase Orders.",
                        ActionUrl = $"/production/maker-orders?orderNo={orderNo}",
                        Priority = 2, // High
                        DueDate = DateTime.Today.AddDays(3),
                        CreatedBy = plannerUser,
                        UpdateExistingIfExists = true
                    });

                    _logger.LogInformation("Purchaser role task dispatched for Order #{OrderNo} ({PendingPurchaseQty} pending pcs)", orderNo, totalPendingPurchaseQty);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to create Purchaser role task for Order #{OrderNo}", orderNo);
                }
            }

            // ── 3. Hub Supervisors Tasks (Active/Unlocked Items with In-House Production Only) ──
            foreach (var item in request.Items.Where(i => !i.IsLockedInUi && (i.IsPlanned || i.HasDeltaQty)))
            {
                // Only dispatch hub tasks if manufacturing quantity is allocated
                if (item.ProductionQty <= 0)
                {
                    continue;
                }

                foreach (var hub in item.HubSchedules)
                {
                    var targetUserNames = hub.Supervisors
                        .Where(s => !string.IsNullOrWhiteSpace(s.UserName))
                        .Select(s => s.UserName.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    if (!targetUserNames.Any())
                    {
                        continue;
                    }

                    try
                    {
                        int plannedQty = hub.PlannedQty > 0 ? hub.PlannedQty : item.ProductionQty;

                        await _workflowTaskEngine.CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                        {
                            SourceEntityType = "ProductionHub",
                            SourceEntityRefId = $"{orderNo}_{item.CompItemID}_Hub{hub.Hub_Name}",
                            TargetRole = "HubSupervisor",
                            TargetUserNames = targetUserNames,
                            Title = $"Production Task: {item.CompItemID} - Hub {hub.Hub_Name} (Order #{orderNo})",
                            Description = $"Production scheduled for item {item.CompItemID} ({item.ItemName}) at Hub {hub.Hub_Name} for Order #{orderNo}. Planned Quantity: {plannedQty:N0} pcs. Schedule: {hub.StartDate:dd-MMM-yyyy} to {hub.EndDate:dd-MMM-yyyy}.",
                            ActionUrl = $"/production/process-groups",
                            Priority = 1, // Medium
                            DueDate = hub.EndDate, // Due date as per End Date selected
                            CreatedBy = plannerUser
                        });

                        _logger.LogInformation("Hub task dispatched for Item {ItemID} Hub {Hub} targeting {UserCount} supervisors with due date {DueDate:dd-MMM-yyyy}",
                            item.CompItemID, hub.Hub_Name, targetUserNames.Count, hub.EndDate);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Failed to create Hub task for Item {ItemID} Hub {Hub} in Order #{OrderNo}",
                            item.CompItemID, hub.Hub_Name, orderNo);
                    }
                }
            }
        }
    }
}
