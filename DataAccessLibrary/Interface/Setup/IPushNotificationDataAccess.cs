using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace DataAccessLibrary.Interface.Setup
{
    public interface IPushNotificationDataAccess
    {
        Task UpsertSubscriptionAsync(string userName, int? userId, string endpoint, string p256dh, string auth, string? deviceName);
        Task<List<UserPushSubscriptionModel>> GetSubscriptionsByUserNameAsync(string userName);
        Task<List<UserPushSubscriptionModel>> GetSubscriptionsByUserIdAsync(string userId);
        Task<List<UserPushSubscriptionModel>> GetAllActiveSubscriptionsAsync();
        Task DeleteSubscriptionByEndpointAsync(string endpoint);
        Task DeactivateSubscriptionAsync(string endpoint);
    }
}
