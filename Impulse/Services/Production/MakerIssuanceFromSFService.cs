using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerIssuanceFromSFService : IMakerIssuanceFromSFService
    {
        private readonly IMakerIssuanceFromSFDataAccess _dataAccess;

        public MakerIssuanceFromSFService(IMakerIssuanceFromSFDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ProcessPOLookupModel>> GetAllProcessesAsync()
        {
            return _dataAccess.GetAllProcessesAsync();
        }

        public Task<SFLotLookupResultModel?> LookupLotNoAsync(string lotNo)
        {
            return _dataAccess.LookupLotNoAsync(lotNo);
        }

        public Task<List<SFStockItemLookupModel>> GetAvailableSFItemsAsync(int processId = 0, long vendId = 0, int returnProcessId = 0)
        {
            return _dataAccess.GetAvailableSFItemsAsync(processId, vendId, returnProcessId);
        }

        public Task<List<SFStockOpeningLocationModel>> GetStockOpeningLocationsForItemAsync(string itemId, int processId)
        {
            return _dataAccess.GetStockOpeningLocationsForItemAsync(itemId, processId);
        }

        public Task<List<UnshippedOrderLookupModel>> GetUnshippedOrdersForItemAsync(string itemId)
        {
            return _dataAccess.GetUnshippedOrdersForItemAsync(itemId);
        }

        public Task<long> SaveSFIssuanceAsync(CreateSFIssuanceHeaderModel header, List<CreateSFIssuanceLineModel> lines, string userName, int userId, string machineName)
        {
            return _dataAccess.SaveSFIssuanceAsync(header, lines, userName, userId, machineName);
        }
    }
}
