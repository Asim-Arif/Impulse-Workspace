using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;

[AllowAnonymous]
[IgnoreAntiforgeryToken]
public class LogoutModel : PageModel
{
    private readonly SignInManager<IdentityUser> _signInManager;
    private readonly ILogger<LogoutModel> _logger;
    public string ReturnUrl { get; set; }
    public LogoutModel(SignInManager<IdentityUser> signInManager, ILogger<LogoutModel> logger)
    {
        _signInManager = signInManager;
        _logger = logger;
    }
    public async Task OnGetAsync(string returnUrl = null)
    {
        returnUrl ??= Url.Content("~/"); 
        await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
        ReturnUrl = returnUrl;
    }
    public async Task<IActionResult> OnPostAsync()
    {
        
        await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
        _logger.LogInformation("User logged out.");
        return RedirectToPage("Login");
    }
}
