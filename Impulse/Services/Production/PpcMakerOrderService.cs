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
    public class PpcMakerOrderService : IPpcMakerOrderService
    {
        private readonly IPpcMakerOrderDataAccess _dataAccess;
        private readonly IWorkflowTaskEngine _workflowTaskEngine;
        private readonly ILogger<PpcMakerOrderService> _logger;

        public PpcMakerOrderService(
            IPpcMakerOrderDataAccess dataAccess,
            IWorkflowTaskEngine workflowTaskEngine,
            ILogger<PpcMakerOrderService> logger)
        {
            _dataAccess = dataAccess;
            _workflowTaskEngine = workflowTaskEngine;
            _logger = logger;
        }

        private bool _schemaEnsured = false;
        private async Task EnsureSchemaOnceAsync()
        {
            if (!_schemaEnsured)
            {
                await _dataAccess.EnsureSchemaAsync();
                _schemaEnsured = true;
            }
        }

        public async Task<PpcMakerPoOrderHeaderDto?> GetOrderHeaderAsync(string orderNo)
        {
            await EnsureSchemaOnceAsync();
            return await _dataAccess.GetOrderHeaderAsync(orderNo);
        }

        public async Task<List<PpcMakerPoOrderHeaderDto>> GetActiveOrdersWithPendingPurchasesAsync()
        {
            await EnsureSchemaOnceAsync();
            return await _dataAccess.GetActiveOrdersWithPendingPurchasesAsync();
        }

        public async Task<List<PpcMakerPoItemRowDto>> GetPpcPurchasesForOrderAsync(string orderNo)
        {
            await EnsureSchemaOnceAsync();
            return await _dataAccess.GetPpcPurchasesForOrderAsync(orderNo);
        }

        public async Task<List<MakerPOLookupModel>> GetAllMakersAsync()
        {
            await EnsureSchemaOnceAsync();
            return await _dataAccess.GetAllMakersAsync();
        }

        public async Task<GenerateMakerPoResult> GenerateMakerPosAsync(GenerateMakerPoRequest request)
        {
            var result = await _dataAccess.GenerateMakerPosAsync(request);

            if (result.Success && result.TotalRowsCreated > 0)
            {
                try
                {
                    // Check if all planned purchase items for this order are fulfilled
                    var allLines = await _dataAccess.GetPpcPurchasesForOrderAsync(request.OrderNo);
                    var remainingPending = allLines.Count(l => !l.IsPosted);

                    if (remainingPending == 0)
                    {
                        var poListStr = string.Join(", ", result.GeneratedMasterPoNumbers);
                        await _workflowTaskEngine.CompleteRoleTaskAsync(
                            "CustomerOrder",
                            request.OrderNo,
                            "Purchaser",
                            request.UserName,
                            $"All planned Maker POs generated successfully. Master PO(s): {poListStr}"
                        );
                        _logger.LogInformation("Purchaser task completed for Order #{OrderNo} following generation of POs {PoList}",
                            request.OrderNo, poListStr);
                    }
                    else
                    {
                        _logger.LogInformation("Order #{OrderNo} still has {Count} pending purchase lines remaining after generating {GeneratedCount} POs.",
                            request.OrderNo, remainingPending, result.TotalRowsCreated);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to complete workflow task for Purchaser on Order #{OrderNo}", request.OrderNo);
                }
            }

            return result;
        }
    }
}
