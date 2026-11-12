using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Impulse.Services.IntraOffice;
using Impulse.Services.Notifications;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Security
{
    public class FactoryIpSecurityService : IFactoryIpSecurityService
    {
        private readonly IConfiguration _configuration;
        private readonly IRemoteAccessDataAccess _remoteAccessData;
        private readonly IUserDataAccess _userData;
        private readonly IAppNotificationService _appNotification;
        private readonly IWebPushNotificationService _webPush;
        private readonly IWhatsAppNotificationService _whatsApp;
        private readonly ILogger<FactoryIpSecurityService> _logger;

        private readonly bool _enabled;
        private readonly bool _allowLoopback;
        private readonly List<string> _allowedIps = new();
        private readonly List<(IPAddress Network, int PrefixLength)> _allowedSubnets = new();

        public FactoryIpSecurityService(
            IConfiguration configuration,
            IRemoteAccessDataAccess remoteAccessData,
            IUserDataAccess userData,
            IAppNotificationService appNotification,
            IWebPushNotificationService webPush,
            IWhatsAppNotificationService whatsApp,
            ILogger<FactoryIpSecurityService> logger)
        {
            _configuration = configuration;
            _remoteAccessData = remoteAccessData;
            _userData = userData;
            _appNotification = appNotification;
            _webPush = webPush;
            _whatsApp = whatsApp;
            _logger = logger;

            var section = _configuration.GetSection("SecuritySettings:FactoryIPRestriction");
            _enabled = section.GetValue<bool>("Enabled", false);
            _allowLoopback = section.GetValue<bool>("AllowLoopback", true);

            if (_enabled)
            {
                var ips = section.GetSection("AllowedIPs").Get<List<string>>() ?? new List<string>();
                _allowedIps = ips.Where(ip => !string.IsNullOrWhiteSpace(ip)).Select(ip => ip.Trim()).ToList();

                var subnets = section.GetSection("AllowedSubnets").Get<List<string>>() ?? new List<string>();
                foreach (var subnet in subnets)
                {
                    if (TryParseCidr(subnet, out var network, out var prefix))
                    {
                        _allowedSubnets.Add((network, prefix));
                    }
                }
            }
        }

        public bool IsIpRestrictionEnabled => _enabled;

        public string GetClientIp(HttpContext? context)
        {
            if (context == null) return "127.0.0.1";

            // Check X-Forwarded-For if reverse proxy / Cloudflare / ARR is used
            var forwardedHeader = context.Request.Headers["X-Forwarded-For"].FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(forwardedHeader))
            {
                var firstIp = forwardedHeader.Split(',')[0].Trim();
                if (IPAddress.TryParse(firstIp, out _))
                {
                    return firstIp;
                }
            }

            var remoteIp = context.Connection.RemoteIpAddress;
            if (remoteIp != null)
            {
                // Map IPv6 mapped IPv4 to standard IPv4
                if (remoteIp.IsIPv4MappedToIPv6)
                {
                    remoteIp = remoteIp.MapToIPv4();
                }
                return remoteIp.ToString();
            }

            return "127.0.0.1";
        }

        public bool IsAllowedFactoryIp(string clientIp)
        {
            if (!_enabled) return true;
            if (string.IsNullOrWhiteSpace(clientIp)) return false;

            // Allow local loopback if enabled in settings
            if (_allowLoopback && (clientIp == "127.0.0.1" || clientIp == "::1" || clientIp.Equals("localhost", StringComparison.OrdinalIgnoreCase)))
                return true;

            if (!IPAddress.TryParse(clientIp, out var clientAddress))
                return false;

            if (clientAddress.IsIPv4MappedToIPv6)
                clientAddress = clientAddress.MapToIPv4();

            var normalizedClient = clientAddress.ToString();

            // Check exact IP match
            if (_allowedIps.Any(ip => string.Equals(ip, normalizedClient, StringComparison.OrdinalIgnoreCase)))
                return true;

            // Check CIDR subnets
            foreach (var (network, prefix) in _allowedSubnets)
            {
                if (IsInSubnet(clientAddress, network, prefix))
                    return true;
            }

            return false;
        }

        public async Task<bool> CanUserAccessRemotelyAsync(UserModel user, string clientIp)
        {
            if (!_enabled) return true;

            // 1. If connecting from factory premises, allow immediately
            if (IsAllowedFactoryIp(clientIp))
                return true;

            // 2. If user is explicitly granted permanent remote access, allow
            if (user.AllowOutsideAccess)
                return true;

            // 3. Check active temporary approval
            var activeApproval = await _remoteAccessData.GetActiveApprovalAsync(user.UserID);
            if (activeApproval != null && activeApproval.ExpiresAt > DateTime.Now)
                return true;

            return false;
        }

        public Task<RemoteAccessRequestModel?> GetActiveApprovalAsync(int userId)
        {
            return _remoteAccessData.GetActiveApprovalAsync(userId);
        }

        public async Task<int> RequestAccessAsync(int userId, string userName, string clientIp, string? reason, string? requestBaseUrl = null)
        {
            // Check if user already has an active pending request from the last 15 minutes
            var existing = await _remoteAccessData.GetLatestPendingRequestAsync(userId);
            if (existing != null)
            {
                return existing.EntryID;
            }

            var request = new RemoteAccessRequestModel
            {
                UserID = userId,
                UserName = userName,
                ClientIP = clientIp,
                Reason = string.IsNullOrWhiteSpace(reason) ? "Work from home / remote access requested" : reason.Trim(),
                RequestDT = DateTime.Now,
                Status = "Pending"
            };

            var entryId = await _remoteAccessData.CreateRequestAsync(request);
            request.EntryID = entryId;

            // Dispatch admin notifications in background
            _ = Task.Run(async () =>
            {
                try
                {
                    await DispatchAdminAlertsAsync(request, requestBaseUrl);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to dispatch admin alerts for remote access request {EntryID}", entryId);
                }
            });

            return entryId;
        }

        public async Task<bool> DispatchAdminAlertsAsync(RemoteAccessRequestModel request, string? requestBaseUrl = null)
        {
            try
            {
                var allUsers = await _userData.GetUsersAsync();
                var admins = allUsers.Where(u => 
                    !u.InActive && 
                    (u.UserManagement || 
                     u.AssignedRoles.Any(r => r.Equals("Admin", StringComparison.OrdinalIgnoreCase) || 
                                              r.Equals("Administrator", StringComparison.OrdinalIgnoreCase)) ||
                     u.UserName.Equals("Administrator", StringComparison.OrdinalIgnoreCase) ||
                     u.UserName.Equals("Admin", StringComparison.OrdinalIgnoreCase))
                ).ToList();

                if (!admins.Any())
                {
                    _logger.LogWarning("No active administrator found to receive remote access alerts.");
                    return false;
                }

                var title = "Remote Access Request";
                var body = $"{request.UserName} ({request.EmployeeName ?? "User"}) requested temporary remote access from IP {request.ClientIP}.";

                foreach (var admin in admins)
                {
                    // 1. In-App Notification (Live Toast / Alert in UI)
                    var appNotif = new AppNotification
                    {
                        Category = NotificationCategory.System,
                        Title = title,
                        Message = body,
                        SenderName = request.UserName,
                        TargetUserId = admin.UserName,
                        ActionUrl = "/setup/remote-access",
                        CreatedAt = DateTime.UtcNow
                    };
                    await _appNotification.SendNotificationAsync(appNotif);

                    // 2. Browser Web Push Notification (Desktop & Mobile)
                    try
                    {
                        await _webPush.SendDirectPushAsync(admin.UserName, title, body, "/setup/remote-access");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogDebug("WebPush could not be delivered to {Admin}: {Msg}", admin.UserName, ex.Message);
                    }

                    // 3. WhatsApp Notification
                    if (!string.IsNullOrWhiteSpace(admin.CellNo))
                    {
                        try
                        {
                            var baseAddress = !string.IsNullOrWhiteSpace(requestBaseUrl)
                                ? requestBaseUrl.TrimEnd('/')
                                : (_configuration["AppSettings:BaseAddress"]?.TrimEnd('/') ?? "");
                            if (Uri.TryCreate(baseAddress, UriKind.Absolute, out var parsedUri) &&
                                IPAddress.TryParse(parsedUri.Host, out var ipHost) &&
                                ipHost.AddressFamily == AddressFamily.InterNetwork &&
                                !IPAddress.IsLoopback(ipHost))
                            {
                                var builder = new UriBuilder(parsedUri)
                                {
                                    Host = $"{parsedUri.Host}.sslip.io"
                                };
                                baseAddress = builder.Uri.ToString().TrimEnd('/');
                            }

                            var reviewUrl = string.IsNullOrWhiteSpace(baseAddress)
                                ? "/setup/remote-access"
                                : $"{baseAddress}/setup/remote-access";

                            var waMessage = $"*Impulse ERP - Remote Access Request*\n\n" +
                                            $"User: *{request.UserName}* ({request.EmployeeName ?? "Staff"})\n" +
                                            $"IP Address: *{request.ClientIP}*\n" +
                                            $"Reason: {request.Reason ?? "Not specified"}\n" +
                                            $"Time: {request.RequestDT:dd-MMM-yyyy hh:mm tt}\n\n" +
                                            $"Review & Approve:\n{reviewUrl}";

                            await _whatsApp.SendDirectMessageAsync(admin.CellNo, waMessage);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "WhatsApp alert failed for admin cell {CellNo}", admin.CellNo);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error dispatching admin alerts for remote access request {EntryID}", request.EntryID);
                return false;
            }
        }

        private static bool TryParseCidr(string cidr, out IPAddress network, out int prefixLength)
        {
            network = IPAddress.Any;
            prefixLength = 0;

            if (string.IsNullOrWhiteSpace(cidr)) return false;

            var parts = cidr.Split('/');
            if (parts.Length != 2) return false;

            if (!IPAddress.TryParse(parts[0].Trim(), out network)) return false;
            if (!int.TryParse(parts[1].Trim(), out prefixLength)) return false;

            if (network.AddressFamily == AddressFamily.InterNetwork && (prefixLength < 0 || prefixLength > 32))
                return false;

            if (network.AddressFamily == AddressFamily.InterNetworkV6 && (prefixLength < 0 || prefixLength > 128))
                return false;

            return true;
        }

        private static bool IsInSubnet(IPAddress address, IPAddress network, int prefixLength)
        {
            if (address.AddressFamily != network.AddressFamily)
                return false;

            var addressBytes = address.GetAddressBytes();
            var networkBytes = network.GetAddressBytes();

            int fullBytes = prefixLength / 8;
            int remainingBits = prefixLength % 8;

            for (int i = 0; i < fullBytes; i++)
            {
                if (addressBytes[i] != networkBytes[i])
                    return false;
            }

            if (remainingBits > 0 && fullBytes < addressBytes.Length)
            {
                byte mask = (byte)(0xFF << (8 - remainingBits));
                if ((addressBytes[fullBytes] & mask) != (networkBytes[fullBytes] & mask))
                    return false;
            }

            return true;
        }
    }
}

