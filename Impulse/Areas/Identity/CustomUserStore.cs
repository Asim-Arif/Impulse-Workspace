using Azure.Core.Cryptography;
using Impulse;
using Impulse.Data;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using Microsoft.VisualStudio.Web.CodeGenerators.Mvc.Templates.BlazorIdentity.Pages.Manage;

namespace Impulse.Areas.Identity
{
    public class CustomUserStore
    {
        private readonly ApplicationDbContext _context;
        private readonly UserManager<IdentityUser> _userManager;
        private Encryption encryption=new Encryption(); 
        public CustomUserStore(ApplicationDbContext context, UserManager<IdentityUser> userManager)
        {
            _context = context;
            _userManager = userManager;
        }

        public async Task<User?> GetUserByUsernameAsync(string username)
        {
            if (string.IsNullOrWhiteSpace(username)) return null;
            var trimmed = username.Trim();
            return await _context.Users.FirstOrDefaultAsync(u => u.UserName.ToLower() == trimmed.ToLower());
        }

        public async Task<bool> ValidateCredentialsAsync(string username, string password)
        {
            if (string.IsNullOrWhiteSpace(username) || password == null) return false;
            var trimmed = username.Trim();
            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName.ToLower() == trimmed.ToLower());
            if (user == null || user.InActive == true)
            {
                return false;
            }

            string dbPassword = (user.Password ?? "").Trim();
            return string.Equals(dbPassword, password.Trim(), StringComparison.Ordinal);
        }
    }
}
