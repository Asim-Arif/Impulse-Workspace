using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class ReceiveLotService : IReceiveLotService
    {
        private readonly IReceiveLotDataAccess _dataAccess;

        public ReceiveLotService(IReceiveLotDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<LotSearchResultModel?> SearchLotAsync(string lotNo)
        {
            return _dataAccess.SearchLotAsync(lotNo);
        }

        public Task<List<MasterPOLineItemModel>> GetLotLinesAsync(string lotNo, long vendIssuedEntryId)
        {
            return _dataAccess.GetLotLinesAsync(lotNo, vendIssuedEntryId);
        }

        public Task<long> SaveLotReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName)
        {
            return _dataAccess.SaveLotReceivingAsync(header, lines, userName, userId, machineName);
        }
    }
}
