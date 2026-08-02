using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IMakerPODataAccess
    {
        Task<List<ProcessPOLookupModel>> GetPurchaseProcessesAsync();
        Task<List<MakerPOLookupModel>> GetMakersForProcessAsync(int processId);
        Task<List<AssignedItemLookupModel>> GetAssignedItemsForMakerAndProcessAsync(long vendId, int processId);
        Task<List<IssuableOrderItemModel>> GetIssuableOrderItemsAsync(int processId, string itemId);
        Task<List<EmployeeLookupModel>> GetEmployeesAsync();
        Task<List<SteelTypeLookupModel>> GetSteelTypesAsync();
        Task<List<MakerBalanceSummaryModel>> GetMakerStockBalancesAsync(long vendId);
        Task<List<MakerHistorySummaryModel>> GetMakerIssuanceHistoryAsync(long vendId);
        Task<string> GetNextMasterPONoAsync();
        Task<string> GetNextBatchNoAsync();
        Task<long> SaveMakerPOAsync(CreateMakerPOHeaderModel header, List<CreateMakerPOLineModel> lines, string userName, int userId, string machineName);
    }
}
