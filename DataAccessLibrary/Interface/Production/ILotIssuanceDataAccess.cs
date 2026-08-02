using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface ILotIssuanceDataAccess
    {
        Task<List<ProcessPOLookupModel>> GetProcessesAsync();
        Task<List<MakerPOLookupModel>> GetMakersForProcessAsync(int processId);
        Task<LotIssuanceLookupResultModel> LookupLotForIssuanceAsync(string lotNo);
        Task<List<AvailableLotIssuanceItemModel>> GetAvailableIssuanceItemsAsync(int processId, long vendId);
        Task<long> SaveLotIssuanceAsync(CreateLotIssuanceHeaderModel header, List<CreateLotIssuanceLineModel> lines, string userName, int userId, string machineName);
    }
}
