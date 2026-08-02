using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerPOService : IMakerPOService
    {
        private readonly IMakerPODataAccess _dataAccess;

        public MakerPOService(IMakerPODataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ProcessPOLookupModel>> GetPurchaseProcessesAsync()
        {
            return _dataAccess.GetPurchaseProcessesAsync();
        }

        public Task<List<MakerPOLookupModel>> GetMakersForProcessAsync(int processId)
        {
            return _dataAccess.GetMakersForProcessAsync(processId);
        }

        public Task<List<AssignedItemLookupModel>> GetAssignedItemsForMakerAndProcessAsync(long vendId, int processId)
        {
            return _dataAccess.GetAssignedItemsForMakerAndProcessAsync(vendId, processId);
        }

        public Task<List<IssuableOrderItemModel>> GetIssuableOrderItemsAsync(int processId, string itemId)
        {
            return _dataAccess.GetIssuableOrderItemsAsync(processId, itemId);
        }

        public Task<List<EmployeeLookupModel>> GetEmployeesAsync()
        {
            return _dataAccess.GetEmployeesAsync();
        }

        public Task<List<SteelTypeLookupModel>> GetSteelTypesAsync()
        {
            return _dataAccess.GetSteelTypesAsync();
        }

        public Task<List<MakerBalanceSummaryModel>> GetMakerStockBalancesAsync(long vendId)
        {
            return _dataAccess.GetMakerStockBalancesAsync(vendId);
        }

        public Task<List<MakerHistorySummaryModel>> GetMakerIssuanceHistoryAsync(long vendId)
        {
            return _dataAccess.GetMakerIssuanceHistoryAsync(vendId);
        }

        public Task<string> GetNextMasterPONoAsync()
        {
            return _dataAccess.GetNextMasterPONoAsync();
        }

        public Task<string> GetNextBatchNoAsync()
        {
            return _dataAccess.GetNextBatchNoAsync();
        }

        public Task<long> SaveMakerPOAsync(CreateMakerPOHeaderModel header, List<CreateMakerPOLineModel> lines, string userName, int userId, string machineName)
        {
            return _dataAccess.SaveMakerPOAsync(header, lines, userName, userId, machineName);
        }
    }
}
