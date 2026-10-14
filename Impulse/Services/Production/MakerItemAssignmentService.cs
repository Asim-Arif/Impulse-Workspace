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

        public Task<bool> HasCustomerWiseRatesSupportAsync()
        {
            return _dataAccess.HasCustomerWiseRatesSupportAsync();
        }

        public Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            return _dataAccess.GetCustomersAsync();
        }

        public Task<List<MakerLookupModel>> GetActiveMakersAsync()
        {
            return _dataAccess.GetActiveMakersAsync();
        }

        public Task<List<ProcessLookupModel>> GetMakerAssignedProcessesAsync(long vendId)
        {
            return _dataAccess.GetMakerAssignedProcessesAsync(vendId);
        }

        public Task<List<UnassignedItemModel>> GetUnassignedItemsAsync(long vendId, int processId, string? custCode = null)
        {
            return _dataAccess.GetUnassignedItemsAsync(vendId, processId, custCode);
        }

        public Task<List<AssignedMakerItemModel>> GetAssignedItemsAsync(long vendId, int processId, string? custCode = null)
        {
            return _dataAccess.GetAssignedItemsAsync(vendId, processId, custCode);
        }

        public Task AssignItemAsync(long vendId, int processId, string itemId, decimal rate, string? custCode = null)
        {
            return _dataAccess.AssignItemAsync(vendId, processId, itemId, rate, custCode);
        }

        public Task UpdateAssignedItemRateAndRemarksAsync(long entryId, decimal newRate, decimal oldRate, string newRemarks, string userName)
        {
            return _dataAccess.UpdateAssignedItemRateAndRemarksAsync(entryId, newRate, oldRate, newRemarks, userName);
        }

        public Task UpdateRateForAllCustomersAsync(long vendId, int processId, string itemId, decimal newRate, decimal oldRate, string newRemarks, string userName)
        {
            return _dataAccess.UpdateRateForAllCustomersAsync(vendId, processId, itemId, newRate, oldRate, newRemarks, userName);
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
