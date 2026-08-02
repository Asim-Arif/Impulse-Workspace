using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IReceiveLotService
    {
        Task<LotSearchResultModel?> SearchLotAsync(string lotNo);
        Task<List<MasterPOLineItemModel>> GetLotLinesAsync(string lotNo, long vendIssuedEntryId);
        Task<long> SaveLotReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName);
    }
}
