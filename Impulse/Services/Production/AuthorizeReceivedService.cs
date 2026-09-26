using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Production
{
    public class AuthorizeReceivedService : IAuthorizeReceivedService
    {
        private readonly IAuthorizeReceivedDataAccess _dataAccess;
        private readonly IHubWorkflowOrchestrator _hubOrchestrator;
        private readonly ILogger<AuthorizeReceivedService> _logger;

        public AuthorizeReceivedService(
            IAuthorizeReceivedDataAccess dataAccess,
            IHubWorkflowOrchestrator hubOrchestrator,
            ILogger<AuthorizeReceivedService> logger)
        {
            _dataAccess = dataAccess;
            _hubOrchestrator = hubOrchestrator;
            _logger = logger;
        }

        public Task<List<PendingAuthorizeReceivedItemModel>> GetPendingAuthorizeReceivedItemsAsync(AuthorizeReceivedFilterModel filter)
        {
            return _dataAccess.GetPendingAuthorizeReceivedItemsAsync(filter);
        }

        public Task<List<WastageTypeLookupModel>> GetWastageTypesAsync()
        {
            return _dataAccess.GetWastageTypesAsync();
        }

        public Task<List<RepairTypeLookupModel>> GetRepairTypesAsync()
        {
            return _dataAccess.GetRepairTypesAsync();
        }

        public Task<List<EmployeeLookupModel>> GetInspectorsAsync()
        {
            return _dataAccess.GetInspectorsAsync();
        }

        public Task<List<ProcessPOLookupModel>> GetProcessesAsync()
        {
            return _dataAccess.GetProcessesAsync();
        }

        public Task<List<MakerPOLookupModel>> GetMakersAsync()
        {
            return _dataAccess.GetMakersAsync();
        }

        public async Task<bool> SaveAuthorizationAsync(List<PendingAuthorizeReceivedItemModel> items, string inspectorEmpId, string userName, string machineName, int userId)
        {
            bool success = await _dataAccess.SaveAuthorizationAsync(items, inspectorEmpId, userName, machineName, userId);

            if (success && items != null && items.Any())
            {
                try
                {
                    var distinctAuthorizedLots = items
                        .Where(i => i.IsChecked && !string.IsNullOrWhiteSpace(i.LotNo))
                        .GroupBy(i => new { i.LotNo, i.ProcessID, i.ItemCode })
                        .Select(g => new
                        {
                            LotNo = g.Key.LotNo,
                            ProcessID = g.Key.ProcessID,
                            ItemCode = g.Key.ItemCode,
                            OrderNo = g.FirstOrDefault()?.OrderNo ?? "",
                            TotalRcvdQty = g.Sum(x => x.RcvdQty)
                        });

                    foreach (var lot in distinctAuthorizedLots)
                    {
                        await _hubOrchestrator.HandleLotReceivingHubTransitionAsync(
                            lot.LotNo,
                            lot.ProcessID,
                            lot.ItemCode,
                            lot.OrderNo,
                            lot.TotalRcvdQty,
                            userName);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error executing Hub transition workflow after SaveAuthorizationAsync for User [{UserName}]", userName);
                }
            }

            return success;
        }
    }
}
