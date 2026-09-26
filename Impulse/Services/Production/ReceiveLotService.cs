using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Production
{
    public class ReceiveLotService : IReceiveLotService
    {
        private readonly IReceiveLotDataAccess _dataAccess;
        private readonly IHubWorkflowOrchestrator _hubOrchestrator;
        private readonly ILogger<ReceiveLotService> _logger;

        public ReceiveLotService(
            IReceiveLotDataAccess dataAccess,
            IHubWorkflowOrchestrator hubOrchestrator,
            ILogger<ReceiveLotService> logger)
        {
            _dataAccess = dataAccess;
            _hubOrchestrator = hubOrchestrator;
            _logger = logger;
        }

        public Task<LotSearchResultModel?> SearchLotAsync(string lotNo)
        {
            return _dataAccess.SearchLotAsync(lotNo);
        }

        public Task<List<MasterPOLineItemModel>> GetLotLinesAsync(string lotNo, long vendIssuedEntryId)
        {
            return _dataAccess.GetLotLinesAsync(lotNo, vendIssuedEntryId);
        }

        public async Task<long> SaveLotReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName)
        {
            long receivingHeaderId = await _dataAccess.SaveLotReceivingAsync(header, lines, userName, userId, machineName);

            if (receivingHeaderId > 0 && lines != null && lines.Any())
            {
                try
                {
                    var distinctLots = lines
                        .Where(l => !string.IsNullOrWhiteSpace(l.LotNo) && l.RcvdQty > 0)
                        .GroupBy(l => new { l.LotNo, l.ProcessID, l.ItemCode })
                        .Select(g => new
                        {
                            LotNo = g.Key.LotNo,
                            ProcessID = g.Key.ProcessID,
                            ItemCode = g.Key.ItemCode,
                            OrderNo = g.FirstOrDefault()?.OrderNo ?? "",
                            TotalRcvdQty = g.Sum(x => x.RcvdQty)
                        });

                    foreach (var lot in distinctLots)
                    {
                        bool reqAuth = await _hubOrchestrator.IsAuthRequiredAsync(lot.ProcessID);
                        if (!reqAuth)
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
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error executing Hub transition workflow after SaveLotReceivingAsync for User [{UserName}]", userName);
                }
            }

            return receivingHeaderId;
        }
    }
}
