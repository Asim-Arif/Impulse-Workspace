using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IProductionDeletionDataAccess
    {
        Task<int> CreateRequestAsync(ProductionDeletionRequestModel model);
        Task<List<ProductionDeletionRequestModel>> GetPendingRequestsAsync(string? requestType = null);
        Task<ProductionDeletionRequestModel?> GetRequestByIdAsync(int id);
        Task<ProductionDeletionRequestModel?> GetPendingRequestByEntityIdAsync(string requestType, long entityRefId);
        Task<ProductionDeletionRequestModel?> GetPendingRequestByLotNoAsync(string lotNo);
        Task<bool> IsLotLockedAsync(string lotNo, long vrdEntryId = 0);
        Task<List<long>> GetActivePendingVRDEntryIdsAsync();
        Task<List<long>> GetActivePendingIssuanceEntryIdsAsync();
        Task<List<string>> GetActivePendingSkipProcessLotNosAsync();
        Task<List<string>> GetActivePendingDeletionLotNosAsync();
        Task<bool> IsIssuanceLockedAsync(long issuanceEntryId, string? lotNo = null);
        Task<bool> ApproveRequestAsync(int id, string directorUserName, string? remarks = null);
        Task<bool> RejectRequestAsync(int id, string directorUserName, string remarks);
        Task<bool> UpdateTaskIdAsync(int requestId, int taskId);
    }
}
