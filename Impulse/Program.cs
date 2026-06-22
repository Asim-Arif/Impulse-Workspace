//using Blazored.Toast;
using Impulse.Pages;
using DataAccessLibrary;
using DataAccessLibrary.DAC;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Impulse.Areas.Identity;
using Impulse.Configurations;
using Impulse.Data;
using Impulse.Services;
using Impulse.Services.Integrations;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Options;
using Radzen;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");

builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(connectionString));

builder.Services.AddDatabaseDeveloperPageExceptionFilter();

builder.Services.AddDefaultIdentity<IdentityUser>(options => options.SignIn.RequireConfirmedAccount = true)
    .AddEntityFrameworkStores<ApplicationDbContext>();

builder.Services.AddAuthorization();
builder.Services.AddCascadingAuthenticationState();
builder.Services.AddRazorPages();

builder.Services.AddTransient<IItemData, ItemData>();
builder.Services.AddTransient<ICustData, CustomersData>();

//builder.Services.AddTransient<IDBHelperService, DBHelper>();
builder.Services.AddScoped<IDBHelper,DBHelper>();


builder.Services.AddScoped<IChartOfAccountsDataAccess, ChartOfAccountsDataAccess>();
builder.Services.AddScoped<ICIPData, ChartOfAccounts_DAC>();
builder.Services.AddScoped<IPhotocopyData, PhotoCopyDataAccess>();
builder.Services.AddScoped<IVouchersDataAccess, VouchersDataAccess>();
builder.Services.AddScoped<ICommonDataAccess, CommonDataAccess>();
builder.Services.AddSingleton<UserSessionService>();
builder.Services.AddScoped<UserSessionService>();
builder.Services.AddScoped<IBankListDataAccess, BankListDataAccess>();
builder.Services.AddScoped<IAccountReportingAccess, AccountsReportingDataAccess>();

//builder.Services.AddServerSideBlazor();
builder.Services.AddServerSideBlazor(options =>
{
    options.DetailedErrors = true;
});
builder.Services.AddScoped<AuthenticationStateProvider, RevalidatingIdentityAuthenticationStateProvider<IdentityUser>>();
builder.Services.AddScoped<SignInManager<IdentityUser>>(); 
builder.Services.AddScoped<UserManager<IdentityUser>>();

// Register your custom user store
builder.Services.AddScoped<CustomUserStore>();
builder.Services.AddSingleton<WeatherForecastService>();
builder.Services.AddHttpContextAccessor();
//builder.Services.AddBlazorContextMenu();
builder.Services.AddBlazorContextMenu(options =>
{
    // Define a template named "Win11Style"
    options.ConfigureTemplate("Win11Style", template =>
    {
        // Apply your custom classes here
        template.MenuCssClass = "win11-context-menu-container";
        template.MenuItemCssClass = "win11-menu-item";
        template.SeperatorCssClass = "win11-divider";
        
        // You can also override other parts like the icon wrapper or submenu class if needed
    });
});
builder.Services.AddBlazorBootstrap();
builder.Services.AddRadzenComponents();
builder.Services.AddScoped<Radzen.NotificationService>();

builder.Services.AddScoped<IVoucherService, VoucherService>();
builder.Services.AddScoped<ICommonServices, CommonServices>();
builder.Services.AddScoped<IDBHelperService, DBHelperService>();
builder.Services.AddScoped<IAuditService, AuditService>();
builder.Services.AddScoped<INotificationService,NotificationServiceManager>();
builder.Services.AddScoped<IReportNavigationServiceCA, ReportNavigationServiceCA>();
builder.Services.AddScoped<IReportNavigationService, ReportNavigationService>();
builder.Services.AddScoped<IAccountReportingAccessService, AccountReportingAccessService>();
builder.Services.AddScoped<IFinancialStatementService, FinancialStatementService>();
builder.Services.AddScoped<IChqBookDetail, ChqBookDetail>();
builder.Services.AddScoped<ITrailBalance_BreakupService, TrailBalance_BreakupService>();

builder.Services.AddScoped<IFinancialStatementDataAccess, FinancialStatementDataAccess>();
builder.Services.AddScoped<ITrailBalance_BreakupDataAccess, TrailBalance_BreakupDataAccess>();
builder.Services.AddScoped<IChqBookDetailDataAccess, ChqBookDetailDataAccess>();


builder.Services.AddScoped<IReportManagerService, ReportManagerService>();


builder.Services.AddScoped<LoadingService>();
builder.Services.AddScoped<SecurityService>();

builder.Services.AddHttpClient<IMeezanBankService, MeezanBankService>(client => {
    client.Timeout = TimeSpan.FromSeconds(30);
});

builder.Services.AddHttpClient<IFaysalBankServices, FaysalBankServices>(client => {
    client.Timeout = TimeSpan.FromSeconds(60); // Faysal might need more time for large file generations
});

builder.Services.AddHttpClient("MyApiClient", client => 
{
    var baseAddress = builder.Configuration.GetValue<string>("AppSettings:BaseAddress");
    client.BaseAddress = new Uri(baseAddress); 
});

builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {

        options.Cookie.Name = "auth_token";
        options.LoginPath = "/login";
        options.Cookie.MaxAge = TimeSpan.FromMinutes(30);
        options.AccessDeniedPath = "/login";
    });

/*builder.Services.AddControllers()
    .AddJsonOptions(options => {
        options.JsonSerializerOptions.PropertyNamingPolicy = null; // Preserve casing
    });
*/
//builder.Services.Configure<AppSettings>(builder.Configuration);
builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));
builder.Services.AddSingleton(sp => sp.GetRequiredService<IOptions<AppSettings>>().Value);

var app = builder.Build();
// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseMigrationsEndPoint();
}
else
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();
app.MapFallbackToPage("/_Host");

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.UseEndpoints(endpoints =>
{
    endpoints.MapRazorPages(); // Additional endpoint mappings...
});

app.MapControllers();
app.MapBlazorHub();
app.MapControllers();

app.Use(async (context, next) =>
{
    Console.WriteLine($"Incoming request: {context.Request.Path}");
    await next.Invoke();
});

app.Run();
