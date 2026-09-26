using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Production
{
    public class ReceiveAgainstPOService : IReceiveAgainstPOService
    {
        private readonly IReceiveAgainstPODataAccess _dataAccess;
        private readonly IHubWorkflowOrchestrator _hubOrchestrator;
        private readonly ILogger<ReceiveAgainstPOService> _logger;

        public ReceiveAgainstPOService(
            IReceiveAgainstPODataAccess dataAccess,
            IHubWorkflowOrchestrator hubOrchestrator,
            ILogger<ReceiveAgainstPOService> logger)
        {
            _dataAccess = dataAccess;
            _hubOrchestrator = hubOrchestrator;
            _logger = logger;
        }

        public Task<MasterPOSearchResultModel?> SearchMasterPOAsync(string receiptId)
        {
            return _dataAccess.SearchMasterPOAsync(receiptId);
        }

        public Task<List<MasterPOLineItemModel>> GetMasterPOLinesAsync(long vendIssuedEntryId)
        {
            return _dataAccess.GetMasterPOLinesAsync(vendIssuedEntryId);
        }

        public async Task<long> SavePOReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName)
        {
            long receivingHeaderId = await _dataAccess.SavePOReceivingAsync(header, lines, userName, userId, machineName);

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
                    _logger.LogError(ex, "Error executing Hub transition workflow after SavePOReceivingAsync for User [{UserName}]", userName);
                }
            }

            return receivingHeaderId;
        }
    }
}
