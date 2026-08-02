using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IReceiveLotDataAccess
    {
        Task<LotSearchResultModel?> SearchLotAsync(string lotNo);
        Task<List<MasterPOLineItemModel>> GetLotLinesAsync(string lotNo, long vendIssuedEntryId);
        Task<long> SaveLotReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName);
    }
}
