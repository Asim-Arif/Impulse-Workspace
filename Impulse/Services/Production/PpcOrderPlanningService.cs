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

        private async Task DispatchDownstreamWorkflowTasksAsync(PpcPlanSaveRequest request)
        {
            var orderNo = request.OrderNo;
            var plannerUser = request.PlannedBy ?? "PPC Planner";

            // ── 1. Stock User Role Task & Notification ───────────────────────
            int totalStockQty = request.Items.Sum(i => i.StockQty);
            int stockItemsCount = request.Items.Count(i => i.StockQty > 0);

            if (totalStockQty > 0)
            {
                try
                {
                    await _workflowTaskEngine.CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                    {
                        SourceEntityType = "CustomerOrder",
                        SourceEntityRefId = orderNo,
                        TargetRole = "Stock",
                        Title = $"Stock Issuance Request for Order #{orderNo}",
                        Description = $"PPC has planned in-hand stock issuance of {totalStockQty:N0} pcs across {stockItemsCount} items for Customer Order #{orderNo}. Please review and issue from store.",
                        ActionUrl = $"/stock/stock-order-adjustment?orderNo={orderNo}",
                        Priority = 2, // High
                        DueDate = DateTime.Today.AddDays(2),
                        CreatedBy = plannerUser
                    });

                    _logger.LogInformation("Stock issuance task dispatched for Order #{OrderNo} ({StockQty} pcs)", orderNo, totalStockQty);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to create Stock role task for Order #{OrderNo}", orderNo);
                }
            }

            // ── 2. Purchaser User Role Task & Notification ───────────────────
            int totalPurchaseQty = request.Items.Sum(i => i.TotalPurchaseQty);
            int purchaseItemsCount = request.Items.Count(i => i.TotalPurchaseQty > 0);

            if (totalPurchaseQty > 0)
            {
                try
                {
                    var purchaseSummaryList = new List<string>();
                    foreach (var item in request.Items.Where(i => i.Purchases.Any(p => p.PurchaseQty > 0)))
                    {
                        foreach (var p in item.Purchases.Where(p => p.PurchaseQty > 0))
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
                        Description = $"PPC has planned Maker POs for {totalPurchaseQty:N0} pcs across {purchaseItemsCount} items for Customer Order #{orderNo}.\n{summaryText}\nPlease generate Maker Purchase Orders.",
                        ActionUrl = $"/production/maker-orders?orderNo={orderNo}",
                        Priority = 2, // High
                        DueDate = DateTime.Today.AddDays(3),
                        CreatedBy = plannerUser
                    });

                    _logger.LogInformation("Purchaser role task dispatched for Order #{OrderNo} ({PurchaseQty} pcs)", orderNo, totalPurchaseQty);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to create Purchaser role task for Order #{OrderNo}", orderNo);
                }
            }

            // ── 3. Hub Supervisors Tasks (Per Item & Hub) ─────────────────────
            foreach (var item in request.Items)
            {
                // Only dispatch hub tasks if manufacturing quantity is allocated
                if (item.ProductionQty <= 0 && !item.HubSchedules.Any(h => h.PlannedQty > 0))
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
