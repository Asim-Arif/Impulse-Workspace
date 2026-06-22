using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Http;
using System.Net;

namespace Impulse.Services
{
    public class AuditService : IAuditService
    {
        private readonly AuthenticationStateProvider _authStateProvider;

        
        private readonly IHttpContextAccessor _httpContextAccessor;

        public AuditService(IHttpContextAccessor httpContextAccessor)
        {
            _httpContextAccessor = httpContextAccessor;
        }

        public string GetClientIpAddress()
        {
            var httpContext = _httpContextAccessor.HttpContext;
            if (httpContext == null)
                return "Unknown";
            // ... (Your IP retrieval logic from the previous answer) ...
            var ip = httpContext.Request.Headers["X-Forwarded-For"].FirstOrDefault();
            if (string.IsNullOrEmpty(ip))
            {
                ip = httpContext.Connection.RemoteIpAddress?.ToString();
            }
            return ip ?? "Unknown";
        }

        public string GetCurrentUserName()
        {
            // Requires user to be authenticated
            return _httpContextAccessor.HttpContext?.User?.Identity?.Name ?? "Guest";
        }
    }
}
