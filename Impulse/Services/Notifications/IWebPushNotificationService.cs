using System.Threading.Tasks;
using Impulse.Services.IntraOffice;

namespace Impulse.Services.Notifications
{
    public interface IWebPushNotificationService
    {
        string GetPublicKey();
        Task SendNotificationAsync(AppNotification notification);
        Task SendDirectPushAsync(string userId, string title, string body, string? actionUrl = null, string? icon = null);
        Task SendBroadcastPushAsync(string title, string body, string? actionUrl = null, string? icon = null);
    }
}
