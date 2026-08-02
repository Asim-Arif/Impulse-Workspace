using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerItemAssignmentService : IMakerItemAssignmentService
    {
        private readonly IMakerItemAssignmentDataAccess _dataAccess;

        public MakerItemAssignmentService(IMakerItemAssignmentDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<MakerLookupModel>> GetActiveMakersAsync()
        {
            return _dataAccess.GetActiveMakersAsync();
        }

        public Task<List<ProcessLookupModel>> GetMakerAssignedProcessesAsync(long vendId)
        {
            return _dataAccess.GetMakerAssignedProcessesAsync(vendId);
        }

        public Task<List<UnassignedItemModel>> GetUnassignedItemsAsync(long vendId, int processId)
        {
            return _dataAccess.GetUnassignedItemsAsync(vendId, processId);
        }

        public Task<List<AssignedMakerItemModel>> GetAssignedItemsAsync(long vendId, int processId)
        {
            return _dataAccess.GetAssignedItemsAsync(vendId, processId);
        }

        public Task AssignItemAsync(long vendId, int processId, string itemId, decimal rate)
        {
            return _dataAccess.AssignItemAsync(vendId, processId, itemId, rate);
        }

        public Task UpdateAssignedItemRateAndRemarksAsync(long entryId, decimal newRate, decimal oldRate, string newRemarks, string userName)
        {
            return _dataAccess.UpdateAssignedItemRateAndRemarksAsync(entryId, newRate, oldRate, newRemarks, userName);
        }

        public Task<bool> IsItemReceivableAsync(long vendId, int processId, string itemId)
        {
            return _dataAccess.IsItemReceivableAsync(vendId, processId, itemId);
        }

        public Task UnassignItemAsync(long entryId)
        {
            return _dataAccess.UnassignItemAsync(entryId);
        }

        public Task<List<ItemRevisionHistoryModel>> GetItemRevisionHistoryAsync(long entryId)
        {
            return _dataAccess.GetItemRevisionHistoryAsync(entryId);
        }
    }
}
