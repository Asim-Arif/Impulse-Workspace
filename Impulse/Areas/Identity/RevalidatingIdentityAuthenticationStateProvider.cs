using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Server;
using Microsoft.AspNetCore.Identity;
using Microsoft.Extensions.Options;
using System.Security.Claims;

namespace Impulse.Areas.Identity
{
    public class RevalidatingIdentityAuthenticationStateProvider<TUser>
        : RevalidatingServerAuthenticationStateProvider where TUser : class
    {
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly IdentityOptions _options;

        public RevalidatingIdentityAuthenticationStateProvider(
            ILoggerFactory loggerFactory,
            IServiceScopeFactory scopeFactory,
            IOptions<IdentityOptions> optionsAccessor)
            : base(loggerFactory)
        {
            _scopeFactory = scopeFactory;
            _options = optionsAccessor.Value;
        }

        protected override TimeSpan RevalidationInterval => TimeSpan.FromMinutes(30);

        protected override async Task<bool> ValidateAuthenticationStateAsync(
            AuthenticationState authenticationState, CancellationToken cancellationToken)
        {
            var user = authenticationState?.User;
            if (user?.Identity?.IsAuthenticated != true)
            {
                return false;
            }

            string? username = user.Identity.Name;
            if (string.IsNullOrWhiteSpace(username))
            {
                return false;
            }

            var scope = _scopeFactory.CreateScope();
            try
            {
                var userStore = scope.ServiceProvider.GetRequiredService<CustomUserStore>();
                var dbUser = await userStore.GetUserByUsernameAsync(username);
                return dbUser != null && dbUser.InActive != true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (scope is IAsyncDisposable asyncDisposable)
                {
                    await asyncDisposable.DisposeAsync();
                }
                else
                {
                    scope.Dispose();
                }
            }
        }
    }
}
