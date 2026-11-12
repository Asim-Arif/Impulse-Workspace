using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;
using Microsoft.AspNetCore.Http;

namespace Impulse.Services.Security
{
    public interface IFactoryIpSecurityService
    {
        bool IsIpRestrictionEnabled { get; }
        string GetClientIp(HttpContext? context);
        bool IsAllowedFactoryIp(string clientIp);
        Task<bool> CanUserAccessRemotelyAsync(UserModel user, string clientIp);
        Task<RemoteAccessRequestModel?> GetActiveApprovalAsync(int userId);
        Task<int> RequestAccessAsync(int userId, string userName, string clientIp, string? reason, string? requestBaseUrl = null);
        Task<bool> DispatchAdminAlertsAsync(RemoteAccessRequestModel request, string? requestBaseUrl = null);
    }
}

