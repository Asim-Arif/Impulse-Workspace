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

//Payroll Module Registrations
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.ITakeAttendanceDataAccess, DataAccessLibrary.DAC.Payroll.TakeAttendanceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.ITakeAttendanceService, Impulse.Services.Payroll.TakeAttendanceService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IOverTimeAuthDataAccess, DataAccessLibrary.DAC.Payroll.OverTimeAuthDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IOverTimeAuthService, Impulse.Services.Payroll.OverTimeAuthService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IManualAttendanceDataAccess, DataAccessLibrary.DAC.Payroll.ManualAttendanceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IManualAttendanceService, Impulse.Services.Payroll.ManualAttendanceService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IEmployeeTransferDataAccess, DataAccessLibrary.DAC.Payroll.EmployeeTransferDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IEmployeeTransferService, Impulse.Services.Payroll.EmployeeTransferService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.ISalaryHistoryDataAccess, DataAccessLibrary.DAC.Payroll.SalaryHistoryDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.ISalaryHistoryService, Impulse.Services.Payroll.SalaryHistoryService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IBiometricLedgerDataAccess, DataAccessLibrary.DAC.Payroll.BiometricLedgerDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IBiometricLedgerService, Impulse.Services.Payroll.BiometricLedgerService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IAttendanceLedgerDataAccess, DataAccessLibrary.DAC.Payroll.AttendanceLedgerDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IAttendanceLedgerService, Impulse.Services.Payroll.AttendanceLedgerService>();
builder.Services.AddScoped<Impulse.Services.Payroll.IIFaceMachineService, Impulse.Services.Payroll.IFaceMachineService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IEmployeeLeaveDataAccess, DataAccessLibrary.DAC.Payroll.EmployeeLeaveDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IEmployeeLeaveService, Impulse.Services.Payroll.EmployeeLeaveService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IMonthlyAttendanceDataAccess, DataAccessLibrary.DAC.Payroll.MonthlyAttendanceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IMonthlyAttendanceService, Impulse.Services.Payroll.MonthlyAttendanceService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IGatePassDataAccess, DataAccessLibrary.DAC.Payroll.GatePassDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IGatePassService, Impulse.Services.Payroll.GatePassService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IAdvanceLongDataAccess, DataAccessLibrary.DAC.Payroll.AdvanceLongDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IAdvanceLongService, Impulse.Services.Payroll.AdvanceLongService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IAdvanceShortDataAccess, DataAccessLibrary.DAC.Payroll.AdvanceShortDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IAdvanceShortService, Impulse.Services.Payroll.AdvanceShortService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IEmpToEmpTransferDataAccess, DataAccessLibrary.DAC.Payroll.EmpToEmpTransferDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IEmpToEmpTransferService, Impulse.Services.Payroll.EmpToEmpTransferService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IAdjustDedAmtDataAccess, DataAccessLibrary.DAC.Payroll.AdjustDedAmtDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IAdjustDedAmtService, Impulse.Services.Payroll.AdjustDedAmtService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IAdvanceLedgerDataAccess, DataAccessLibrary.DAC.Payroll.AdvanceLedgerDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IAdvanceLedgerService, Impulse.Services.Payroll.AdvanceLedgerService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IPostFineDataAccess, DataAccessLibrary.DAC.Payroll.PostFineDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IPostFineService, Impulse.Services.Payroll.PostFineService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IEmpFineLedgerDataAccess, DataAccessLibrary.DAC.Payroll.EmpFineLedgerDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IEmpFineLedgerService, Impulse.Services.Payroll.EmpFineLedgerService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.ISalarySheetDataAccess, DataAccessLibrary.DAC.Payroll.SalarySheetDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.ISalarySheetService, Impulse.Services.Payroll.SalarySheetService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IClearAdvanceDataAccess, DataAccessLibrary.DAC.Payroll.ClearAdvanceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IClearAdvanceService, Impulse.Services.Payroll.ClearAdvanceService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IClearShortTermDataAccess, DataAccessLibrary.DAC.Payroll.ClearShortTermDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IClearShortTermService, Impulse.Services.Payroll.ClearShortTermService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.IHoldSalaryDataAccess, DataAccessLibrary.DAC.Payroll.HoldSalaryDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IHoldSalaryService, Impulse.Services.Payroll.HoldSalaryService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.IGratuityDataAccess, DataAccessLibrary.DAC.Payroll.GratuityDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IGratuityService, Impulse.Services.Payroll.GratuityService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.IAbsentSheetDataAccess, DataAccessLibrary.DAC.Payroll.AbsentSheetDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IAbsentSheetService, Impulse.Services.Payroll.AbsentSheetService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.ISocialSecurityDataAccess, DataAccessLibrary.DAC.Payroll.SocialSecurityDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.ISocialSecurityService, Impulse.Services.Payroll.SocialSecurityService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.IEOBIDataAccess, DataAccessLibrary.DAC.Payroll.EOBIDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IEOBIService, Impulse.Services.Payroll.EOBIService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.IDesignationDataAccess, DataAccessLibrary.DAC.Payroll.DesignationDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IDesignationService, Impulse.Services.Payroll.DesignationService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.IPayrollSettingsDataAccess, DataAccessLibrary.DAC.Payroll.PayrollSettingsDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IPayrollSettingsService, Impulse.Services.Payroll.PayrollSettingsService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.IPayrollPoliciesDataAccess, DataAccessLibrary.DAC.Payroll.PayrollPoliciesDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IPayrollPoliciesService, Impulse.Services.Payroll.PayrollPoliciesService>();
builder.Services.AddScoped<DataAccessLibrary.DAC.Payroll.IHolidayDataAccess, DataAccessLibrary.DAC.Payroll.HolidayDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IHolidayService, Impulse.Services.Payroll.HolidayService>();

//Stock Module Registrations
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IFinishMovementDataAccess, DataAccessLibrary.DAC.Stock.FinishMovementDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IFinishMovementService, Impulse.Services.Stock.FinishMovementService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IFinishItemLedgerDataAccess, DataAccessLibrary.DAC.Stock.FinishItemLedgerDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IFinishItemLedgerService, Impulse.Services.Stock.FinishItemLedgerService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IFinishTransactionDataAccess, DataAccessLibrary.DAC.Stock.FinishTransactionDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IFinishTransactionService, Impulse.Services.Stock.FinishTransactionService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IFinishStockReceivingDataAccess, DataAccessLibrary.DAC.Stock.FinishStockReceivingDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IFinishStockReceivingService, Impulse.Services.Stock.FinishStockReceivingService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IFinishStockIssuanceDataAccess, DataAccessLibrary.DAC.Stock.FinishStockIssuanceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IFinishStockIssuanceService, Impulse.Services.Stock.FinishStockIssuanceService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.ISFTransactionDataAccess, DataAccessLibrary.DAC.Stock.SFTransactionDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.ISFTransactionService, Impulse.Services.Stock.SFTransactionService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.ISFMovementDataAccess, DataAccessLibrary.DAC.Stock.SFMovementDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.ISFMovementService, Impulse.Services.Stock.SFMovementService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.ISemiFinishOpenReceivingDataAccess, DataAccessLibrary.DAC.Stock.SemiFinishOpenReceivingDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.ISemiFinishOpenReceivingService, Impulse.Services.Stock.SemiFinishOpenReceivingService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IStockLedgerDataAccess, DataAccessLibrary.DAC.Stock.StockLedgerDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IStockLedgerService, Impulse.Services.Stock.StockLedgerService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IChangeBatchNoLotDataAccess, DataAccessLibrary.DAC.Stock.ChangeBatchNoLotDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IChangeBatchNoLotService, Impulse.Services.Stock.ChangeBatchNoLotService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IRMMovementDataAccess, DataAccessLibrary.DAC.Stock.RMMovementDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IRMMovementService, Impulse.Services.Stock.RMMovementService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendorRMAssignmentDataAccess, DataAccessLibrary.DAC.Stock.VendorRMAssignmentDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendorRMAssignmentService, Impulse.Services.Stock.VendorRMAssignmentService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IRMIssuanceDataAccess, DataAccessLibrary.DAC.Stock.RMIssuanceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IRMIssuanceService, Impulse.Services.Stock.RMIssuanceService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IRMIssuanceListDataAccess, DataAccessLibrary.DAC.Stock.RMIssuanceListDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IRMIssuanceListService, Impulse.Services.Stock.RMIssuanceListService>();

builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IRMGroupsDataAccess, DataAccessLibrary.DAC.Stock.RMGroupsDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IRMGroupsService, Impulse.Services.Stock.RMGroupsService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.INewRMDataAccess, DataAccessLibrary.DAC.Stock.NewRMDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.INewRMService, Impulse.Services.Stock.NewRMService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.INewVendorDataAccess, DataAccessLibrary.DAC.Stock.NewVendorDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.INewVendorService, Impulse.Services.Stock.NewVendorService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendOrderDataAccess, DataAccessLibrary.DAC.Stock.VendOrderDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendOrderService, Impulse.Services.Stock.VendOrderService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendGateRcvdDataAccess, DataAccessLibrary.DAC.Stock.VendGateRcvdDataAccess>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendGateRcvdListDataAccess, DataAccessLibrary.DAC.Stock.VendGateRcvdListDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendGateRcvdService, Impulse.Services.Stock.VendGateRcvdService>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendGateRcvdListService, Impulse.Services.Stock.VendGateRcvdListService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendGateRcvdInspectionDataAccess, DataAccessLibrary.DAC.Stock.VendGateRcvdInspectionDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendGateRcvdInspectionService, Impulse.Services.Stock.VendGateRcvdInspectionService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IMaterialPlacementDataAccess, DataAccessLibrary.DAC.Stock.MaterialPlacementDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IMaterialPlacementService, Impulse.Services.Stock.MaterialPlacementService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IRMDataAccess, DataAccessLibrary.DAC.Stock.RMDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IRMService, Impulse.Services.Stock.RMService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IMaterialPlacementListDataAccess, DataAccessLibrary.DAC.Stock.MaterialPlacementListDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IMaterialPlacementListService, Impulse.Services.Stock.MaterialPlacementListService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendorBillingListDataAccess, DataAccessLibrary.DAC.Stock.VendorBillingListDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendorBillingListService, Impulse.Services.Stock.VendorBillingListService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IChangeBatchNoDataAccess, DataAccessLibrary.DAC.Stock.ChangeBatchNoDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IChangeBatchNoService, Impulse.Services.Stock.ChangeBatchNoService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IVendorBillingDataAccess, DataAccessLibrary.DAC.Stock.VendorBillingDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IVendorBillingService, Impulse.Services.Stock.VendorBillingService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Stock.IRMIssuanceDataAccess, DataAccessLibrary.DAC.Stock.RMIssuanceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Stock.IRMIssuanceService, Impulse.Services.Stock.RMIssuanceService>();
builder.Services.AddScoped<IDepartmentDataAccess, DepartmentDataAccess>();
builder.Services.AddScoped<IDepartmentService, DepartmentService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Payroll.IShortTermSheetDataAccess, DataAccessLibrary.DAC.Payroll.ShortTermSheetDataAccess>();
builder.Services.AddScoped<Impulse.Services.Payroll.IShortTermSheetService, Impulse.Services.Payroll.ShortTermSheetService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Accounts.IMakerShortTermLoanDataAccess, DataAccessLibrary.DAC.Accounts.MakerShortTermLoanDataAccess>();
builder.Services.AddScoped<Impulse.Services.Accounts.IMakerShortTermLoanService, Impulse.Services.Accounts.MakerShortTermLoanService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Accounts.IMakerLongTermLoanDataAccess, DataAccessLibrary.DAC.Accounts.MakerLongTermLoanDataAccess>();
builder.Services.AddScoped<Impulse.Services.Accounts.IMakerLongTermLoanService, Impulse.Services.Accounts.MakerLongTermLoanService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Accounts.IMakerShortLoanLedgerDataAccess, DataAccessLibrary.DAC.Accounts.MakerShortLoanLedgerDataAccess>();
builder.Services.AddScoped<Impulse.Services.Accounts.IMakerShortLoanLedgerService, Impulse.Services.Accounts.MakerShortLoanLedgerService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Accounts.IMakerLongLoanLedgerDataAccess, DataAccessLibrary.DAC.Accounts.MakerLongLoanLedgerDataAccess>();
builder.Services.AddScoped<Impulse.Services.Accounts.IMakerLongLoanLedgerService, Impulse.Services.Accounts.MakerLongLoanLedgerService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Accounts.IMakerLoanClearanceDataAccess, DataAccessLibrary.DAC.Accounts.MakerLoanClearanceDataAccess>();
builder.Services.AddScoped<Impulse.Services.Accounts.IMakerLoanClearanceService, Impulse.Services.Accounts.MakerLoanClearanceService>();
builder.Services.AddScoped<DataAccessLibrary.Interface.Accounts.IMakerLoanTransferDataAccess, DataAccessLibrary.DAC.Accounts.MakerLoanTransferDataAccess>();
builder.Services.AddScoped<Impulse.Services.Accounts.IMakerLoanTransferService, Impulse.Services.Accounts.MakerLoanTransferService>();







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
