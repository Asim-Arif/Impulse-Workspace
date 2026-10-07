using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Impulse.Services.IntraOffice;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using WebPush;

namespace Impulse.Services.Notifications
{
    public class WebPushNotificationService : IWebPushNotificationService
    {
        private static VapidDetails? _cachedVapidDetails;
        private static string? _cachedPublicKey;
        private static readonly object _keyLock = new();

        private readonly IPushNotificationDataAccess _pushData;
        private readonly ILogger<WebPushNotificationService> _logger;
        private readonly VapidDetails _vapidDetails;
        private readonly string _publicKey;

        public WebPushNotificationService(
            IPushNotificationDataAccess pushData,
            IConfiguration configuration,
            ILogger<WebPushNotificationService> logger)
        {
            _pushData = pushData;
            _logger = logger;

            lock (_keyLock)
            {
                if (_cachedVapidDetails == null || string.IsNullOrWhiteSpace(_cachedPublicKey))
                {
                    var subject = configuration["VapidKeys:Subject"] ?? "mailto:support@impulseerp.com";
                    var publicKey = configuration["VapidKeys:PublicKey"];
                    var privateKey = configuration["VapidKeys:PrivateKey"];

                    if (string.IsNullOrWhiteSpace(publicKey) || string.IsNullOrWhiteSpace(privateKey))
                    {
                        _logger.LogInformation("Generating static persistent VAPID keys as none were found in configuration.");
                        var generatedKeys = VapidHelper.GenerateVapidKeys();
                        publicKey = generatedKeys.PublicKey;
                        privateKey = generatedKeys.PrivateKey;
                    }

                    _cachedPublicKey = publicKey;
                    _cachedVapidDetails = new VapidDetails(subject, publicKey, privateKey);
                }
            }

            _publicKey = _cachedPublicKey!;
            _vapidDetails = _cachedVapidDetails!;
        }

        public string GetPublicKey() => _publicKey;

        public async Task SendNotificationAsync(AppNotification notification)
        {
            if (notification == null) return;

            var title = string.IsNullOrWhiteSpace(notification.Title) ? "Impulse ERP" : notification.Title;
            var body = notification.Message ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(notification.SenderName) && !title.Contains(notification.SenderName))
            {
                title = $"{notification.SenderName}: {title}";
            }

            var url = notification.ActionUrl;
            if (string.IsNullOrWhiteSpace(url)) url = "/";

            if (!string.IsNullOrWhiteSpace(notification.TargetUserId))
            {
                await SendDirectPushAsync(notification.TargetUserId, title, body, url, "/favicon.ico");
            }
            else
            {
                await SendBroadcastPushAsync(title, body, url, "/favicon.ico");
            }
        }

        public async Task SendDirectPushAsync(string userId, string title, string body, string? actionUrl = null, string? icon = null)
        {
            if (string.IsNullOrWhiteSpace(userId)) return;

            try
            {
                var subscriptions = await _pushData.GetSubscriptionsByUserIdAsync(userId);
                if (subscriptions == null || subscriptions.Count == 0) return;

                var payload = BuildPayload(title, body, actionUrl, icon);
                await DispatchToSubscriptionsAsync(subscriptions, payload);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to send web push notifications to user {UserId}", userId);
            }
        }

        public async Task SendBroadcastPushAsync(string title, string body, string? actionUrl = null, string? icon = null)
        {
            try
            {
                var subscriptions = await _pushData.GetAllActiveSubscriptionsAsync();
                if (subscriptions == null || subscriptions.Count == 0) return;

                var payload = BuildPayload(title, body, actionUrl, icon);
                await DispatchToSubscriptionsAsync(subscriptions, payload);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to broadcast web push notifications");
            }
        }

        private static string BuildPayload(string title, string body, string? actionUrl, string? icon)
        {
            var dataObj = new
            {
                title = title,
                body = body,
                message = body,
                icon = string.IsNullOrWhiteSpace(icon) ? "/favicon.ico" : icon,
                badge = "/favicon.ico",
                data = new
                {
                    url = string.IsNullOrWhiteSpace(actionUrl) ? "/" : actionUrl
                }
            };

            return JsonSerializer.Serialize(dataObj);
        }

        private async Task DispatchToSubscriptionsAsync(IEnumerable<UserPushSubscriptionModel> subscriptions, string payload)
        {
            var client = new WebPushClient();

            foreach (var sub in subscriptions)
            {
                if (string.IsNullOrWhiteSpace(sub.Endpoint) || string.IsNullOrWhiteSpace(sub.P256dh) || string.IsNullOrWhiteSpace(sub.Auth))
                    continue;

                try
                {
                    var pushSub = new PushSubscription(sub.Endpoint, sub.P256dh, sub.Auth);
                    await client.SendNotificationAsync(pushSub, payload, _vapidDetails);
                }
                catch (WebPushException ex)
                {
                    _logger.LogWarning("WebPush error ({StatusCode}) for endpoint {Endpoint}: {Message}", ex.StatusCode, sub.Endpoint, ex.Message);

                    // 404 Not Found or 410 Gone means the subscription has expired or unsubscribed
                    if (ex.StatusCode == HttpStatusCode.Gone || ex.StatusCode == HttpStatusCode.NotFound)
                    {
                        await _pushData.DeactivateSubscriptionAsync(sub.Endpoint);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Unexpected error sending push notification to endpoint {Endpoint}", sub.Endpoint);
                }
            }
        }
    }
}
