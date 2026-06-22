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
    public class LoginModel : PageModel
    {

        private readonly CustomUserStore _userStore;
        private readonly SignInManager<IdentityUser> _signInManager;
        private readonly ILogger<LoginModel> _logger;
        public readonly UserSessionService _usersession;

        public LoginModel(CustomUserStore userStore, SignInManager<IdentityUser> signInManager, ILogger<LoginModel> logger, IDBHelper idbhelper, UserSessionService userSession) 
        { 
            _userStore = userStore; 
            _signInManager = signInManager; 
            _logger = logger;
            _idbhelper = idbhelper;
            _usersession = userSession;
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
                string currentusername = await _idbhelper.getSingleStringValue("UserName", "Users", $"WHERE EmpID ='{Input.UserName}' AND Inactive=0");
                
                if (currentusername == "") 
                {
                    currentusername = Input.UserName;
                }

                if (await _userStore.ValidateCredentialsAsync(currentusername, Input.Password))
                //if (await _userStore.ValidateCredentialsAsync(Input.EmpID, Input.Password))
                    {
                    
                    _usersession.UserName = currentusername;

                    _logger.LogInformation("User logged in.");
                    // Add your logic to sign in the user and set the authentication cookie

                    // Fetch the UserName according to EmpID
                    //string currentusername = await _idbHelperService.getSingleStringValue("UserName", "Users", $"WHERE EmpID = '{Input.EmpID}'");

                    var claims = new List<Claim>
                    {
                        //new Claim(ClaimTypes.Name, Input.UserName),
                        //new Claim(ClaimTypes.Name, Input.EmpID),
                        new Claim(ClaimTypes.Name, currentusername ?? Input.UserName),
                        new Claim(ClaimTypes.Role, "User")

                    };

                    var claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
                    var principal = new ClaimsPrincipal(claimsIdentity);

                    var authProperties = new AuthenticationProperties 
                    {
                        IsPersistent = Input.RememberMe
                    };

                    await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal, authProperties);
                    
                    //await HttpContext.SignInAsync(principal);
                    
                    // Log claims for debugging
                    _logger.LogInformation("User authenticated with claims: {Claims}", string.Join(", ", claims.Select(c => c.Type + ": " + c.Value)));
                    return LocalRedirect(returnUrl);
                    //return RedirectToPage("Login");

                    //_NavigationManager.NavigateTo("/");
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
