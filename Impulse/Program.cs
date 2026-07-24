//using Blazored.Toast;
using Impulse.Pages;
using DataAccessLibrary;
using DataAccessLibrary.DAC;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.DAC.Export;
using DataAccessLibrary.Models;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.DAC.Payroll;
using Impulse.Services.Payroll;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Impulse.Areas.Identity;
using Impulse.Configurations;
using Impulse.Data;
using Impulse.Services;
using Impulse.Services.Export;
using Impulse.Services.Integrations;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.DAC.Export;
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
builder.Services.AddTransient<ICustomerOrderDataAccess, CustomerOrderDataAccess>();
builder.Services.AddTransient<ICustomerOrderService, CustomerOrderService>();
builder.Services.AddTransient<ICustomerItemBalancesDataAccess, CustomerItemBalancesDataAccess>();
builder.Services.AddScoped<ICustomerItemBalancesService, CustomerItemBalancesService>();
builder.Services.AddScoped<IArticlewiseShippedStatusDataAccess, ArticlewiseShippedStatusDataAccess>();
builder.Services.AddScoped<IArticlewiseShippedStatusService, ArticlewiseShippedStatusService>();
builder.Services.AddScoped<IProformaListDataAccess, ProformaListDataAccess>();
builder.Services.AddScoped<IProformaListService, ProformaListService>();

builder.Services.AddScoped<IInvoiceListDataAccess, InvoiceListDataAccess>();
builder.Services.AddScoped<IInvoiceListService, InvoiceListService>();
builder.Services.AddScoped<ICustomInvoiceAuthDataAccess, CustomInvoiceAuthDataAccess>();
builder.Services.AddScoped<ICustomInvoiceAuthService, CustomInvoiceAuthService>();
builder.Services.AddScoped<IEmployeeDataAccess, EmployeeDataAccess>();
builder.Services.AddScoped<IEmployeeService, EmployeeService>();
builder.Services.AddScoped<IAccountReportingAccessService, AccountReportingAccessService>();

//Stock Module Registrations
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendorRMAssignmentDataAccess, DataAccessLibrary.DAC.Stock.VendorRMAssignmentDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendorRMAssignmentService, Impulse.Services.Stock.VendorRMAssignmentService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IRMGroupsDataAccess, DataAccessLibrary.DAC.Stock.RMGroupsDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IRMGroupsService, Impulse.Services.Stock.RMGroupsService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.INewRMDataAccess, DataAccessLibrary.DAC.Stock.NewRMDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.INewRMService, Impulse.Services.Stock.NewRMService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.INewVendorDataAccess, DataAccessLibrary.DAC.Stock.NewVendorDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.INewVendorService, Impulse.Services.Stock.NewVendorService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendOrderDataAccess, DataAccessLibrary.DAC.Stock.VendOrderDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendOrderService, Impulse.Services.Stock.VendOrderService>();
builder.Services.AddScoped<IDepartmentDataAccess, DepartmentDataAccess>();
builder.Services.AddScoped<IDepartmentService, DepartmentService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Company.IItemDataAccess, DataAccessLibrary.DAC.Company.ItemDataAccess>();
builder.Services.AddScoped<Impulse.Services.Company.IItemService, Impulse.Services.Company.ItemService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Company.ICustomerCatalogDataAccess, DataAccessLibrary.DAC.Company.CustomerCatalogDataAccess>();
builder.Services.AddScoped<Impulse.Services.Company.ICustomerCatalogService, Impulse.Services.Company.CustomerCatalogService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Company.IItemGroupsDataAccess, DataAccessLibrary.DAC.Company.ItemGroupsDataAccess>();
builder.Services.AddScoped<Impulse.Services.Company.IItemGroupsService, Impulse.Services.Company.ItemGroupsService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Company.IMiscSetupDataAccess, DataAccessLibrary.DAC.Company.MiscSetupDataAccess>();
builder.Services.AddScoped<Impulse.Services.Company.IMiscSetupService, Impulse.Services.Company.MiscSetupService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Company.ILocationsDataAccess, DataAccessLibrary.DAC.Company.LocationsDataAccess>();
builder.Services.AddScoped<Impulse.Services.Company.ILocationsService, Impulse.Services.Company.LocationsService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Company.ICurrencyExchangeRatesDataAccess, DataAccessLibrary.DAC.Company.CurrencyExchangeRatesDataAccess>();
builder.Services.AddScoped<Impulse.Services.Company.ICurrencyExchangeRatesService, Impulse.Services.Company.CurrencyExchangeRatesService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Company.IStoresDataAccess, DataAccessLibrary.DAC.Company.StoresDataAccess>();
builder.Services.AddScoped<Impulse.Services.Company.IStoresService, Impulse.Services.Company.StoresService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Company.ISteelDataAccess, DataAccessLibrary.DAC.Company.SteelDataAccess>();
builder.Services.AddScoped<Impulse.Services.Company.ISteelService, Impulse.Services.Company.SteelService>();
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

builder.Services.AddScoped<IForeignCustomerDataAccess, ForeignCustomerDataAccess>();
builder.Services.AddScoped<Impulse.Services.IForeignCustomerService, Impulse.Services.ForeignCustomerService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.ICustomerOrderDataAccess, DataAccessLibrary.DAC.Export.CustomerOrderDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.ICustomerOrderService, Impulse.Services.Export.CustomerOrderService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.ICustomerQuotationDataAccess, DataAccessLibrary.DAC.Export.CustomerQuotationDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.ICustomerQuotationService, Impulse.Services.Export.CustomerQuotationService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IAdvancePaymentDataAccess, DataAccessLibrary.DAC.Export.AdvancePaymentDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IAdvancePaymentService, Impulse.Services.Export.AdvancePaymentService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IOrderItemListDataAccess, DataAccessLibrary.DAC.Export.OrderItemListDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IOrderItemListService, Impulse.Services.Export.OrderItemListService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IProformaInvoiceDataAccess, DataAccessLibrary.DAC.Export.ProformaInvoiceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IProformaInvoiceService, Impulse.Services.Export.ProformaInvoiceService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.ICustomInvoiceDataAccess, DataAccessLibrary.DAC.Export.CustomInvoiceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.ICustomInvoiceService, Impulse.Services.Export.CustomInvoiceService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.ICustomPaymentDataAccess, DataAccessLibrary.DAC.Export.CustomPaymentDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.ICustomPaymentService, Impulse.Services.Export.CustomPaymentService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IPackingLabelsDataAccess, DataAccessLibrary.DAC.Export.PackingLabelsDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IPackingLabelsService, Impulse.Services.Export.PackingLabelsService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IPrintInnerLabelsDataAccess, DataAccessLibrary.DAC.Export.PrintInnerLabelsDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IPrintInnerLabelsService, Impulse.Services.Export.PrintInnerLabelsService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IShippingInstructionsDataAccess, DataAccessLibrary.DAC.Export.ShippingInstructionsDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IShippingInstructionsService, Impulse.Services.Export.ShippingInstructionsService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IPrintValuationDataAccess, DataAccessLibrary.DAC.Export.PrintValuationDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IPrintValuationService, Impulse.Services.Export.PrintValuationService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.ICommercialCoveringDataAccess, DataAccessLibrary.DAC.Export.CommercialCoveringDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.ICommercialCoveringService, Impulse.Services.Export.CommercialCoveringService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.ITotalExportDataAccess, DataAccessLibrary.DAC.Export.TotalExportDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.ITotalExportService, Impulse.Services.Export.TotalExportService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.ITotalDemandDataAccess, DataAccessLibrary.DAC.Export.TotalDemandDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.ITotalDemandService, Impulse.Services.Export.TotalDemandService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IArticlewiseSaleDataAccess, DataAccessLibrary.DAC.Export.ArticlewiseSaleDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IArticlewiseSaleService, Impulse.Services.Export.ArticlewiseSaleService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IArticlewiseDemandDataAccess, DataAccessLibrary.DAC.Export.ArticlewiseDemandDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IArticlewiseDemandService, Impulse.Services.Export.ArticlewiseDemandService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Export.IExportPerformanceReportDataAccess, DataAccessLibrary.DAC.Export.ExportPerformanceReportDataAccess>();
builder.Services.AddScoped<Impulse.Services.Export.IExportPerformanceReportService, Impulse.Services.Export.ExportPerformanceReportService>();

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
