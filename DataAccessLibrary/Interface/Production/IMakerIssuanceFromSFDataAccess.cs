using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IMakerIssuanceFromSFDataAccess
    {
        Task<List<ProcessPOLookupModel>> GetAllProcessesAsync();
        Task<SFLotLookupResultModel?> LookupLotNoAsync(string lotNo);
        Task<List<SFStockItemLookupModel>> GetAvailableSFItemsAsync(int processId = 0, long vendId = 0, int returnProcessId = 0);
        Task<List<SFStockOpeningLocationModel>> GetStockOpeningLocationsForItemAsync(string itemId, int processId);
        Task<List<UnshippedOrderLookupModel>> GetUnshippedOrdersForItemAsync(string itemId);
        Task<long> SaveSFIssuanceAsync(CreateSFIssuanceHeaderModel header, List<CreateSFIssuanceLineModel> lines, string userName, int userId, string machineName);
    }
}
