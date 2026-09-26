using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.WorkflowTasks;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Stock
{
    public class PpcStockAdjustmentService : IPpcStockAdjustmentService
    {
        private readonly IPpcStockAdjustmentDataAccess _dataAccess;
        private readonly IWorkflowTaskEngine _workflowTaskEngine;
        private readonly ILogger<PpcStockAdjustmentService> _logger;

        public PpcStockAdjustmentService(
            IPpcStockAdjustmentDataAccess dataAccess,
            IWorkflowTaskEngine workflowTaskEngine,
            ILogger<PpcStockAdjustmentService> logger)
        {
            _dataAccess = dataAccess;
            _workflowTaskEngine = workflowTaskEngine;
            _logger = logger;
        }

        public async Task<PpcStockAdjustmentOrderHeaderModel?> GetOrderHeaderAsync(string orderNo)
        {
            return await _dataAccess.GetOrderHeaderAsync(orderNo);
        }

        public async Task<List<PpcStockAdjustmentOrderHeaderModel>> GetActiveOrdersWithStockPlanningAsync()
        {
            return await _dataAccess.GetActiveOrdersWithStockPlanningAsync();
        }

        public async Task<List<PpcStockAdjustmentPlannedItemModel>> GetPlannedStockItemsAsync(string orderNo)
        {
            var items = await _dataAccess.GetPlannedStockItemsAsync(orderNo);

            foreach (var item in items)
            {
                item.Processes = await _dataAccess.GetProcessesForItemAsync(item.ItemID);
                if (item.Processes.Any() && !item.SelectedProcessId.HasValue)
                {
                    item.SelectedProcessId = item.Processes.First().ProcessID;
                }

                item.AvailableShelfLots = await _dataAccess.GetAvailableShelfLotsAsync(item.ItemID);

                // Auto-allocate initial quantities
                await AutoAllocateStockAsync(item);
            }

            return items;
        }

        public Task AutoAllocateStockAsync(PpcStockAdjustmentPlannedItemModel item)
        {
            if (item == null) return Task.CompletedTask;

            int targetQty = Math.Min(item.RemainingQtyToIssue, (int)item.AvailableInHandQty);
            item.AdjQty = targetQty;
            item.Allocations = new List<PpcStockShelfAllocationModel>();

            int remainingNeeded = targetQty;
            foreach (var lot in item.AvailableShelfLots.OrderBy(l => l.Batch_No))
            {
                if (remainingNeeded <= 0) break;

                int canTake = Math.Min(remainingNeeded, (int)lot.NetAvailableQty);
                if (canTake > 0)
                {
                    item.Allocations.Add(new PpcStockShelfAllocationModel
                    {
                        Shelf_RefID = lot.Shelf_RefID,
                        StoreName = lot.StoreName,
                        RackNo = lot.RackNo,
                        ShelfNo = lot.ShelfNo,
                        LotNo = lot.LotNo,
                        Batch_No = lot.Batch_No,
                        Mill_Certificate_No = lot.Mill_Certificate_No,
                        RCV_ISD_RefID = lot.RCV_ISD_RefID,
                        NetAvailableQty = lot.NetAvailableQty,
                        AllocatedQty = canTake
                    });
                    remainingNeeded -= canTake;
                }
            }

            return Task.CompletedTask;
        }

        public async Task<PpcStockAdjustmentResultModel> SaveStockAdjustmentAsync(PpcStockAdjustmentSubmissionModel submission)
        {
            var result = await _dataAccess.SaveStockAdjustmentAsync(submission);

            if (result.Success)
            {
                try
                {
                    // Check if all planned stock items for this order are fulfilled
                    var plannedItems = await _dataAccess.GetPlannedStockItemsAsync(submission.OrderNo);
                    var remainingPending = plannedItems.Sum(i => i.RemainingQtyToIssue);

                    if (remainingPending == 0)
                    {
                        await _workflowTaskEngine.CompleteRoleTaskAsync(
                            "CustomerOrder",
                            submission.OrderNo,
                            "Stock",
                            submission.UserName,
                            $"Stock order adjustment completed for Order #{submission.OrderNo}."
                        );
                        _logger.LogInformation("Stock task completed for Order #{OrderNo} following stock adjustment.", submission.OrderNo);
                    }
                    else
                    {
                        _logger.LogInformation("Order #{OrderNo} has {Remaining} remaining stock units to be adjusted.", submission.OrderNo, remainingPending);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to complete workflow task for Stock on Order #{OrderNo}", submission.OrderNo);
                }
            }

            return result;
        }
    }
}
