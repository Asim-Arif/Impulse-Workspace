using System;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using Impulse.Services.Security;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;

namespace Impulse.Middleware
{
    public class RemoteAccessExpirationMiddleware
    {
        private readonly RequestDelegate _next;

        public RemoteAccessExpirationMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        public async Task InvokeAsync(HttpContext context, IFactoryIpSecurityService ipSecurity, IUserDataAccess userData)
        {
            if (ipSecurity.IsIpRestrictionEnabled && context.User?.Identity?.IsAuthenticated == true)
            {
                var path = context.Request.Path.Value ?? string.Empty;

                // Bypass static assets, blazor socket, and authentication endpoints
                bool isAuthOrStatic = path.StartsWith("/Identity/Account/Logout", StringComparison.OrdinalIgnoreCase) ||
                                      path.StartsWith("/Identity/Account/Login", StringComparison.OrdinalIgnoreCase) ||
                                      path.StartsWith("/_blazor", StringComparison.OrdinalIgnoreCase) ||
                                      path.StartsWith("/css", StringComparison.OrdinalIgnoreCase) ||
                                      path.StartsWith("/js", StringComparison.OrdinalIgnoreCase) ||
                                      path.StartsWith("/lib", StringComparison.OrdinalIgnoreCase) ||
                                      path.StartsWith("/images", StringComparison.OrdinalIgnoreCase);

                if (!isAuthOrStatic)
                {
                    var clientIp = ipSecurity.GetClientIp(context);

                    if (!ipSecurity.IsAllowedFactoryIp(clientIp))
                    {
                        var userName = context.User.Identity.Name;
                        if (!string.IsNullOrEmpty(userName))
                        {
                            var user = await userData.GetUserByUserNameAsync(userName);
                            if (user != null && !user.AllowOutsideAccess)
                            {
                                var activeApproval = await ipSecurity.GetActiveApprovalAsync(user.UserID);
                                if (activeApproval == null || activeApproval.ExpiresAt <= DateTime.Now)
                                {
                                    // Remote session has expired or is not active
                                    await context.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
                                    await context.SignOutAsync(IdentityConstants.ApplicationScheme);

                                    context.Response.Redirect("/Identity/Account/Login?remoteExpired=1");
                                    return;
                                }
                            }
                        }
                    }
                }
            }

            await _next(context);
        }
    }
}

