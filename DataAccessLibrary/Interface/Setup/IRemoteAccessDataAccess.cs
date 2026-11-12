using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace DataAccessLibrary.Interface.Setup
{
    public interface IRemoteAccessDataAccess
    {
        Task<int> CreateRequestAsync(RemoteAccessRequestModel request);
        Task<RemoteAccessRequestModel?> GetActiveApprovalAsync(int userId);
        Task<RemoteAccessRequestModel?> GetLatestPendingRequestAsync(int userId);
        Task<RemoteAccessRequestModel?> GetLatestRequestByUserAsync(int userId);
        Task<List<RemoteAccessRequestModel>> GetPendingRequestsAsync();
        Task<List<RemoteAccessRequestModel>> GetRecentRequestsAsync(int limit = 50);
        Task<bool> ApproveRequestAsync(int entryId, int actionByUserId, string actionByUserName, int durationMinutes, string? remarks);
        Task<bool> RejectRequestAsync(int entryId, int actionByUserId, string actionByUserName, string? remarks);
        Task<int> GetPendingCountAsync();
    }
}

