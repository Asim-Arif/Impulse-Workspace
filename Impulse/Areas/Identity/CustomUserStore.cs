using Azure.Core.Cryptography;
using CIP.Data;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using Microsoft.VisualStudio.Web.CodeGenerators.Mvc.Templates.BlazorIdentity.Pages.Manage;

namespace CIP.Areas.Identity
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

        public async Task<bool> ValidateCredentialsAsync(string username, string password)
        {
            var user = await _context.Users.SingleOrDefaultAsync(u => u.UserName == username);
            //var user = await _context.Users.SingleOrDefaultAsync(u => u.EmpID == empid);
            if (user == null)
            {
                return false;
            }

            // Add your password hashing/validation logic here
            //return await _userManager.CheckPasswordAsync(user, password);
            //return user.Password == password;
            //CryptRC4(FromHexDump(strPassword), "awmsial")
            string strPassword="";
            strPassword=encryption.CryptRC4(user.Password, "awmsial");
            //return user.Password ==  password;
            return strPassword == password;

        }
    }
}
