using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class LotIssuanceService : ILotIssuanceService
    {
        private readonly ILotIssuanceDataAccess _dataAccess;

        public LotIssuanceService(ILotIssuanceDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ProcessPOLookupModel>> GetProcessesAsync()
        {
            return _dataAccess.GetProcessesAsync();
        }

        public Task<List<MakerPOLookupModel>> GetMakersForProcessAsync(int processId)
        {
            return _dataAccess.GetMakersForProcessAsync(processId);
        }

        public Task<LotIssuanceLookupResultModel> LookupLotForIssuanceAsync(string lotNo)
        {
            return _dataAccess.LookupLotForIssuanceAsync(lotNo);
        }

        public Task<List<AvailableLotIssuanceItemModel>> GetAvailableIssuanceItemsAsync(int processId, long vendId)
        {
            return _dataAccess.GetAvailableIssuanceItemsAsync(processId, vendId);
        }

        public Task<long> SaveLotIssuanceAsync(CreateLotIssuanceHeaderModel header, List<CreateLotIssuanceLineModel> lines, string userName, int userId, string machineName)
        {
            return _dataAccess.SaveLotIssuanceAsync(header, lines, userName, userId, machineName);
        }
    }
}
