// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
#nullable disable

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.AspNetCore.Authentication.Cookies;
using System.Security.Claims;
using System.Security.Principal;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Components.Routing;
using Microsoft.AspNetCore.Components.Web.Virtualization;
using Microsoft.Identity.Client;
using Microsoft.AspNetCore.Components;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using static System.Runtime.InteropServices.JavaScript.JSType;
using DataAccessLibrary;
using DataAccessLibrary.Models;
using Microsoft.PowerBI.Api.Models;

namespace Impulse.Areas.Identity.Pages.Account
{
    [IgnoreAntiforgeryToken]
    public class LoginModel : PageModel
    {

        private readonly CustomUserStore _userStore;
        private readonly SignInManager<IdentityUser> _signInManager;
        private readonly ILogger<LoginModel> _logger;
        public readonly UserSessionService _usersession;
        private readonly DataAccessLibrary.Interface.Setup.IUserRoleDataAccess _userRoleData;
        private readonly DataAccessLibrary.Interface.Setup.IUserDataAccess _userData;

        public LoginModel(
            CustomUserStore userStore, 
            SignInManager<IdentityUser> signInManager, 
            ILogger<LoginModel> logger, 
            IDBHelper idbhelper, 
            UserSessionService userSession,
            DataAccessLibrary.Interface.Setup.IUserRoleDataAccess userRoleData,
            DataAccessLibrary.Interface.Setup.IUserDataAccess userData) 
        { 
            _userStore = userStore; 
            _signInManager = signInManager; 
            _logger = logger;
            _idbhelper = idbhelper;
            _usersession = userSession;
            _userRoleData = userRoleData;
            _userData = userData;
        }

        [BindProperty]
        public InputModel Input { get; set; }
        public IList<AuthenticationScheme> ExternalLogins { get; set; }
        public string ReturnUrl { get; set; }
        [TempData] public string ErrorMessage { get; set; }
        public class InputModel
        {
            [Required]
            public string UserName { get; set; }
            //public string EmpID { get; set; }

            [Required]
            [DataType(DataType.Password)]
            public string Password { get; set; }

            public bool RememberMe { get; set; }
        }
        public async Task OnGetAsync(string returnUrl = null)
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                ModelState.AddModelError(string.Empty, ErrorMessage);
            }
            returnUrl ??= Url.Content("~/"); // Clear the existing external cookie to ensure a clean login process
            /*await HttpContext.SignOutAsync(IdentityConstants.ExternalScheme);
            ExternalLogins = (await _signInManager.GetExternalAuthenticationSchemesAsync()).ToList();*/
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            ReturnUrl = returnUrl;
        }
        private readonly IDBHelper _idbhelper;

        public async Task<IActionResult> OnPostAsync(string returnUrl = null)
        {
            returnUrl ??= Url.Content("~/");
            
            if (ModelState.IsValid)
            {
                /*string currentusername = await _idbhelper.getSingleStringValue("UserName", "Users", $"WHERE EmpID ='{Input.UserName}' AND Inactive=0");
                
                if (currentusername == "") 
                {
                    currentusername = Input.UserName;
                }*/
                string currentusername = Input.UserName?.Trim();

                var dbUser = await _userStore.GetUserByUsernameAsync(currentusername);
                if (dbUser != null && dbUser.InActive == true)
                {
                    ModelState.AddModelError(string.Empty, "This user account is inactive. Please contact the administrator.");
                    return Page();
                }

                if (await _userStore.ValidateCredentialsAsync(currentusername, Input.Password))
                {
                    currentusername = dbUser?.UserName ?? currentusername;
                    _usersession.UserName = currentusername;

                    _logger.LogInformation("User {UserName} logged in.", currentusername);

                    var claims = new List<Claim>
                    {
                        new Claim(ClaimTypes.Name, currentusername),
                        new Claim(ClaimTypes.NameIdentifier, dbUser != null ? dbUser.UserID.ToString() : currentusername),
                        new Claim(ClaimTypes.Role, "User")
                    };

                    // Load user roles from database
                    if (dbUser != null)
                    {
                        try
                        {
                            var userRoles = await _userRoleData.GetRolesByUserIdAsync(dbUser.UserID);
                            foreach (var role in userRoles.Where(r => !string.IsNullOrWhiteSpace(r)))
                            {
                                if (!claims.Any(c => c.Type == ClaimTypes.Role && c.Value.Equals(role, StringComparison.OrdinalIgnoreCase)))
                                {
                                    claims.Add(new Claim(ClaimTypes.Role, role));
                                }
                            }

                            var fullUser = await _userData.GetUserByUserNameAsync(currentusername);
                            if (fullUser?.UserManagement == true)
                            {
                                if (!claims.Any(c => c.Type == ClaimTypes.Role && c.Value.Equals("Administrator", StringComparison.OrdinalIgnoreCase)))
                                    claims.Add(new Claim(ClaimTypes.Role, "Administrator"));
                                if (!claims.Any(c => c.Type == ClaimTypes.Role && c.Value.Equals("Admin", StringComparison.OrdinalIgnoreCase)))
                                    claims.Add(new Claim(ClaimTypes.Role, "Admin"));
                                if (!claims.Any(c => c.Type == ClaimTypes.Role && c.Value.Equals("Director", StringComparison.OrdinalIgnoreCase)))
                                    claims.Add(new Claim(ClaimTypes.Role, "Director"));
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Could not load database roles for user {UserName}", currentusername);
                        }
                    }

                    if (string.Equals(currentusername, "admin", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(currentusername, "administrator", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!claims.Any(c => c.Type == ClaimTypes.Role && c.Value.Equals("Administrator", StringComparison.OrdinalIgnoreCase)))
                            claims.Add(new Claim(ClaimTypes.Role, "Administrator"));
                        if (!claims.Any(c => c.Type == ClaimTypes.Role && c.Value.Equals("Admin", StringComparison.OrdinalIgnoreCase)))
                            claims.Add(new Claim(ClaimTypes.Role, "Admin"));
                        if (!claims.Any(c => c.Type == ClaimTypes.Role && c.Value.Equals("Director", StringComparison.OrdinalIgnoreCase)))
                            claims.Add(new Claim(ClaimTypes.Role, "Director"));
                    }

                    var claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
                    var principal = new ClaimsPrincipal(claimsIdentity);

                    var authProperties = new AuthenticationProperties 
                    {
                        IsPersistent = Input.RememberMe,
                        ExpiresUtc = DateTimeOffset.UtcNow.AddDays(30)
                    };

                    await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal, authProperties);
                    
                    _logger.LogInformation("User authenticated with claims: {Claims}", string.Join(", ", claims.Select(c => c.Type + ": " + c.Value)));
                    return LocalRedirect(returnUrl);
                }
                else
                {
                    ModelState.AddModelError(string.Empty, "Invalid login attempt.");
                    return Page();
                }
            }

            return Page();
        }
    }

}
