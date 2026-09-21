using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface ILotIssuanceService
    {
        Task<List<ProcessPOLookupModel>> GetProcessesAsync();
        Task<List<MakerPOLookupModel>> GetMakersForProcessAsync(int processId);
        Task<LotIssuanceLookupResultModel> LookupLotForIssuanceAsync(string lotNo);
        Task<List<AvailableLotIssuanceItemModel>> GetAvailableIssuanceItemsAsync(int processId, long vendId);
        Task<long> SaveLotIssuanceAsync(CreateLotIssuanceHeaderModel header, List<CreateLotIssuanceLineModel> lines, string userName, int userId, string machineName);
        Task<List<ProcessPOLookupModel>> GetSubsequentProcessesForSkipAsync(string itemCode, int currentProcessId, bool isReworkLot, int repairType);
        Task<bool> SkipProcessAsync(string itemCode, int currentProcessId, int newProcessId, string lotNo);
    }
}
