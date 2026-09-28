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

                // ── 2. Create Tasks & Notifications for Hub Supervisors (Process -> Hub) ──
                try
                {
                    var generatedLines = request.Lines.Where(l => !string.IsNullOrWhiteSpace(l.MasterPONo)).ToList();
                    if (!generatedLines.Any())
                    {
                        generatedLines = request.Lines.Where(l => l.IsSelected).ToList();
                    }

                    if (generatedLines.Any())
                    {
                        var itemIds = generatedLines.Select(l => l.CompItemID).Distinct();
                        var procIds = generatedLines.Select(l => l.ProcessID).Distinct();
                        var hubMappings = await _dataAccess.GetProcessHubSupervisorsAsync(itemIds, procIds);

                        // Attach Hub info and Supervisors to each generated line
                        var lineHubDetails = generatedLines.Select(line =>
                        {
                            var matchingHubs = hubMappings
                                .Where(m => m.ItemID.Equals(line.CompItemID, StringComparison.OrdinalIgnoreCase) && m.ProcessID == line.ProcessID)
                                .ToList();

                            string hubName = matchingHubs.FirstOrDefault(m => !string.IsNullOrWhiteSpace(m.Hub_Name))?.Hub_Name ?? "Production";
                            string procName = matchingHubs.FirstOrDefault(m => !string.IsNullOrWhiteSpace(m.ProcessName))?.ProcessName ?? line.ProcessName;
                            var supervisorUsers = matchingHubs
                                .Where(m => !string.IsNullOrWhiteSpace(m.UserName))
                                .Select(m => m.UserName!.Trim())
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();

                            string masterPo = !string.IsNullOrWhiteSpace(line.MasterPONo)
                                ? line.MasterPONo
                                : result.GeneratedMasterPoNumbers.FirstOrDefault() ?? "PO";

                            return new
                            {
                                Line = line,
                                MasterPoNo = masterPo,
                                HubName = hubName,
                                ProcessName = procName,
                                Supervisors = supervisorUsers
                            };
                        }).ToList();

                        // Group by (MasterPoNo, HubName) so supervisors get a unified, consolidated notification per Hub
                        var hubGroups = lineHubDetails.GroupBy(x => new { x.MasterPoNo, x.HubName, Maker = x.Line.MakerName });

                        foreach (var group in hubGroups)
                        {
                            var targetUserNames = group.SelectMany(x => x.Supervisors)
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();

                            int totalQty = group.Sum(x => x.Line.PurchaseQty);
                            var minReturnDt = group.Min(x => x.Line.ReturnDT);
                            var summaryItems = string.Join("\n", group.Select(x => $"• {x.Line.CompItemID}: {x.Line.PurchaseQty:N0} pcs at [{x.ProcessName}]"));

                            await _workflowTaskEngine.CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                            {
                                SourceEntityType = "MakerPO",
                                SourceEntityRefId = $"{group.Key.MasterPoNo}_Hub{group.Key.HubName}",
                                TargetRole = "HubSupervisor",
                                TargetUserNames = targetUserNames,
                                Title = $"Maker PO #{group.Key.MasterPoNo} (Hub {group.Key.HubName}): Order #{request.OrderNo}",
                                Description = $"Maker PO #{group.Key.MasterPoNo} ({totalQty:N0} pcs across {group.Count()} item(s)) generated with {group.Key.Maker} for Hub {group.Key.HubName} on Customer Order #{request.OrderNo}.\n{summaryItems}\nExpected Return Date: {minReturnDt:dd-MMM-yyyy}.\nPlease monitor delivery and receiving for your hub station.",
                                ActionUrl = $"/production/maker-orders?orderNo={request.OrderNo}",
                                Priority = 2, // High
                                DueDate = minReturnDt > DateTime.Today ? minReturnDt : DateTime.Today.AddDays(7),
                                CreatedBy = request.UserName
                            });

                            _logger.LogInformation("Hub supervisor task dispatched for Master PO #{PoNo} Hub {Hub} targeting {UserCount} supervisors",
                                group.Key.MasterPoNo, group.Key.HubName, targetUserNames.Count);
                        }
                    }
                }
                catch (Exception hubEx)
                {
                    _logger.LogError(hubEx, "Failed to dispatch Hub supervisor tasks for generated POs on Order #{OrderNo}", request.OrderNo);
                }
            }

            return result;
        }
    }
}
