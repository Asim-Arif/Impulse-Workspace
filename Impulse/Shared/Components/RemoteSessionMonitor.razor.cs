using System;
using System.Security.Claims;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using Impulse.Services.Security;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Routing;
using Microsoft.Extensions.Logging;

namespace Impulse.Shared.Components
{
    public partial class RemoteSessionMonitor : ComponentBase, IDisposable
    {
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private IRemoteAccessDataAccess RemoteAccessData { get; set; } = default!;
        [Inject] private IUserDataAccess UserData { get; set; } = default!;
        [Inject] private IFactoryIpSecurityService IpSecurity { get; set; } = default!;
        [Inject] private ILogger<RemoteSessionMonitor> Logger { get; set; } = default!;

        private PeriodicTimer? _timer;
        private CancellationTokenSource? _cts;
        private int _userId;
        private bool _isRemoteSession;
        private bool _isRedirecting;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                if (!IpSecurity.IsIpRestrictionEnabled)
                    return;

                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                var user = authState.User;

                if (user?.Identity?.IsAuthenticated != true)
                    return;

                var isRemoteClaim = user.FindFirst("IsRemoteSession")?.Value;
                var userIdClaim = user.FindFirst("RemoteUserId")?.Value 
                               ?? user.FindFirst(ClaimTypes.NameIdentifier)?.Value;

                var userName = user.Identity.Name;

                if (string.Equals(isRemoteClaim, "true", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(userIdClaim, out var parsedId))
                {
                    _userId = parsedId;
                    _isRemoteSession = true;
                }
                else if (!string.IsNullOrEmpty(userName))
                {
                    // Fallback for sessions where the claim was not set
                    var dbUser = await UserData.GetUserByUserNameAsync(userName);
                    if (dbUser != null && !dbUser.AllowOutsideAccess)
                    {
                        var latest = await RemoteAccessData.GetLatestRequestByUserAsync(dbUser.UserID);
                        if (latest != null && (latest.Status == "Approved" || latest.Status == "Revoked") && latest.RequestDT > DateTime.Now.AddDays(-1))
                        {
                            _userId = dbUser.UserID;
                            _isRemoteSession = true;
                        }
                    }
                }

                if (_isRemoteSession && _userId > 0)
                {
                    NavigationManager.LocationChanged += OnLocationChanged;

                    // Immediate verification
                    await CheckSessionValidityAsync();

                    // Periodic check every 5 seconds
                    _cts = new CancellationTokenSource();
                    _timer = new PeriodicTimer(TimeSpan.FromSeconds(5));
                    _ = MonitorLoopAsync(_cts.Token);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Error initializing RemoteSessionMonitor");
            }
        }

        private async Task MonitorLoopAsync(CancellationToken token)
        {
            try
            {
                while (_timer != null && await _timer.WaitForNextTickAsync(token))
                {
                    if (token.IsCancellationRequested || _isRedirecting)
                        break;

                    await CheckSessionValidityAsync();
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Error in RemoteSessionMonitor loop");
            }
        }

        private void OnLocationChanged(object? sender, LocationChangedEventArgs e)
        {
            if (_isRemoteSession && !_isRedirecting)
            {
                _ = InvokeAsync(async () => await CheckSessionValidityAsync());
            }
        }

        private async Task CheckSessionValidityAsync()
        {
            if (!_isRemoteSession || _isRedirecting || _userId <= 0)
                return;

            try
            {
                var activeApproval = await RemoteAccessData.GetActiveApprovalAsync(_userId);
                if (activeApproval == null || activeApproval.ExpiresAt <= DateTime.Now)
                {
                    _isRedirecting = true;
                    Dispose();

                    Logger.LogWarning("Remote session expired or revoked for UserID {UserID}. Forcing sign-out.", _userId);

                    var targetUrl = "/Identity/Account/Logout?returnUrl=" + Uri.EscapeDataString("/Identity/Account/Login?remoteExpired=1");
                    NavigationManager.NavigateTo(targetUrl, forceLoad: true);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Error checking remote access validity for UserID {UserID}", _userId);
            }
        }

        public void Dispose()
        {
            try
            {
                NavigationManager.LocationChanged -= OnLocationChanged;
            }
            catch { }

            _cts?.Cancel();
            _cts?.Dispose();
            _cts = null;

            _timer?.Dispose();
            _timer = null;
        }
    }
}

