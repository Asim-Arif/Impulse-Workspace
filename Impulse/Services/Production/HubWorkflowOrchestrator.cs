using System;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Production
{
    public class HubWorkflowOrchestrator : IHubWorkflowOrchestrator
    {
        private readonly IHubWorkflowDataAccess _dataAccess;
        private readonly ILogger<HubWorkflowOrchestrator> _logger;

        public HubWorkflowOrchestrator(IHubWorkflowDataAccess dataAccess, ILogger<HubWorkflowOrchestrator> logger)
        {
            _dataAccess = dataAccess;
            _logger = logger;
        }

        public Task<bool> IsAuthRequiredAsync(int processId)
        {
            return _dataAccess.IsProcessAuthRequiredAsync(processId);
        }

        public async Task<HubTransitionExecutionResult> HandleLotReceivingHubTransitionAsync(
            string lotNo,
            int processId,
            string itemCode,
            string orderNo,
            decimal rcvdQty,
            string actorUserName)
        {
            var result = new HubTransitionExecutionResult();

            if (string.IsNullOrWhiteSpace(lotNo) || string.IsNullOrWhiteSpace(itemCode) || processId <= 0)
            {
                result.Handled = false;
                result.Message = "Invalid lot receiving parameters.";
                return result;
            }

            try
            {
                // 1. Analyze lot's position in its Process Group
                var analysis = await _dataAccess.AnalyzeLotProcessHubAsync(itemCode, processId);
                if (analysis == null)
                {
                    result.Handled = false;
                    result.Message = "Process or item not mapped in ProcessGroupsProcesses.";
                    return result;
                }

                result.CurrentHubName = analysis.CurrentHubName;
                result.IsLastProcessOfHub = analysis.IsLastProcessOfHub;
                result.IsFinalHub = analysis.IsFinalGroupHub;
                result.NextHubName = analysis.NextHubName;

                if (!analysis.IsLastProcessOfHub)
                {
                    // Intermediate process within hub; no hub transition occurs
                    result.Handled = true;
                    result.Message = $"Process {processId} is intermediate within Hub {analysis.CurrentHubName}. No transition needed.";
                    return result;
                }

                // 2. Lot received on LAST process of CurrentHub -> Close current hub task
                int closedCount = await _dataAccess.CloseHubTaskAsync(
                    lotNo,
                    analysis.CurrentHubName,
                    actorUserName,
                    $"Lot #{lotNo} completed final process ({analysis.CurrentProcessName}) of Hub {analysis.CurrentHubName}.");

                result.TasksClosed = closedCount;

                // 3. Branch: Intermediate Hub completed vs Final Hub completed
                if (!analysis.IsFinalGroupHub && !string.IsNullOrWhiteSpace(analysis.NextHubName))
                {
                    // Create task for Next Hub supervisors & dispatch notifications
                    int newTaskId = await _dataAccess.CreateNextHubTaskAndNotificationsAsync(
                        analysis,
                        lotNo,
                        itemCode,
                        orderNo,
                        rcvdQty,
                        actorUserName);

                    result.TasksCreated = (newTaskId > 0) ? 1 : 0;
                    result.NotificationsSent = analysis.NextHubSupervisors.Count;
                    result.Handled = true;
                    result.Message = $"Hub {analysis.CurrentHubName} closed. Task #{newTaskId} and {analysis.NextHubSupervisors.Count} notifications dispatched to Hub {analysis.NextHubName}.";
                }
                else
                {
                    // Final Hub completed -> Dispatch informational notification to role 'Dispatch'
                    int dispatchNotifs = await _dataAccess.SendDispatchFinalLotNotificationAsync(
                        lotNo,
                        itemCode,
                        orderNo,
                        rcvdQty,
                        analysis.CurrentProcessName,
                        actorUserName);

                    result.NotificationsSent = dispatchNotifs;
                    result.Handled = true;
                    result.Message = $"Final Hub {analysis.CurrentHubName} completed. {dispatchNotifs} notification(s) dispatched to role 'Dispatch'.";
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling lot receiving hub transition for Lot [{LotNo}], ProcessID {ProcessID}", lotNo, processId);
                result.Handled = false;
                result.Message = ex.Message;
                return result;
            }
        }
    }
}
