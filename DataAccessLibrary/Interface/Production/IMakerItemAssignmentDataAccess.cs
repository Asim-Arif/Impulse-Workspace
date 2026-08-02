using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IMakerItemAssignmentDataAccess
    {
        Task<List<MakerLookupModel>> GetActiveMakersAsync();
        Task<List<ProcessLookupModel>> GetMakerAssignedProcessesAsync(long vendId);
        Task<List<UnassignedItemModel>> GetUnassignedItemsAsync(long vendId, int processId);
        Task<List<AssignedMakerItemModel>> GetAssignedItemsAsync(long vendId, int processId);
        Task AssignItemAsync(long vendId, int processId, string itemId, decimal rate);
        Task UpdateAssignedItemRateAndRemarksAsync(long entryId, decimal newRate, decimal oldRate, string newRemarks, string userName);
        Task<bool> IsItemReceivableAsync(long vendId, int processId, string itemId);
        Task UnassignItemAsync(long entryId);
        Task<List<ItemRevisionHistoryModel>> GetItemRevisionHistoryAsync(long entryId);
    }
}
