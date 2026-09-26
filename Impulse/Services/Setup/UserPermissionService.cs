using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Setup
{
    public class UserPermissionService : IUserPermissionService
    {
        private readonly IUserPermissionDataAccess _permissionData;
        private readonly AuthenticationStateProvider _authStateProvider;
        private readonly ILogger<UserPermissionService> _logger;

        private bool _isInitialized = false;
        private string _currentUserName = string.Empty;
        private UserModel? _currentUserModel;
        private HashSet<string> _allowedOptionIds = new(StringComparer.OrdinalIgnoreCase);

        public bool IsAdministrator { get; private set; } = false;
        public int CurrentUserId => _currentUserModel?.UserID ?? 0;
        public string CurrentUserName => _currentUserName;

        public UserPermissionService(
            IUserPermissionDataAccess permissionData,
            AuthenticationStateProvider authStateProvider,
            ILogger<UserPermissionService> logger)
        {
            _permissionData = permissionData;
            _authStateProvider = authStateProvider;
            _logger = logger;
        }

        public async Task InitializeAsync(string? userName = null)
        {
            try
            {
                if (string.IsNullOrEmpty(userName))
                {
                    var authState = await _authStateProvider.GetAuthenticationStateAsync();
                    userName = authState.User.Identity?.Name ?? string.Empty;
                }

                // If already initialized for this exact username, return immediately
                if (_isInitialized && string.Equals(_currentUserName, userName, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                _currentUserName = userName;
                _allowedOptionIds.Clear();

                if (string.IsNullOrEmpty(userName))
                {
                    IsAdministrator = false;
                    _currentUserModel = null;
                    _isInitialized = true;
                    return;
                }

                // Administrator check
                if (string.Equals(userName, "Administrator", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(userName, "Admin", StringComparison.OrdinalIgnoreCase))
                {
                    IsAdministrator = true;
                    _currentUserModel = await _permissionData.GetUserPermissionsByUserNameAsync(userName);
                    _isInitialized = true;
                    return;
                }

                // Load User Model from DB
                _currentUserModel = await _permissionData.GetUserPermissionsByUserNameAsync(userName);

                if (_currentUserModel != null && _currentUserModel.UserManagement)
                {
                    IsAdministrator = true;
                    _isInitialized = true;
                    return;
                }

                IsAdministrator = false;

                // Load UserMenuOptions
                _allowedOptionIds = await _permissionData.GetUserMenuOptionIdsByUserNameAsync(userName);
                _isInitialized = true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error initializing user permissions for {User}", userName);
                _isInitialized = true;
            }
        }

        public bool HasModuleAccess(string moduleName)
        {
            if (IsAdministrator) return true;
            if (_currentUserModel == null) return false;

            return moduleName.ToLowerInvariant() switch
            {
                "financial" or "accounts" => _currentUserModel.FinancialMainLink,
                "payroll" or "hr" => _currentUserModel.PayrollMainLink,
                "stock" or "inventory" => _currentUserModel.StockMainLink,
                "production" => _currentUserModel.ProductionMainLink,
                "export" => _currentUserModel.ExportMainLink,
                "company" => _currentUserModel.CompanyMainLink,
                "dashboard" => _currentUserModel.DashBoardMainLink,
                "qms" => _currentUserModel.QMSMainLink,
                "fixedassets" => _currentUserModel.FixedAssetsMainLink,
                "sampling" => _currentUserModel.SamplingMainLink,
                "help" => _currentUserModel.HelpMainLink,
                "office" or "intraoffice" => _currentUserModel.IntraOfficeMainLink,
                _ => true
            };
        }

        private static readonly Dictionary<string, string[]> OptionAliases = new(StringComparer.OrdinalIgnoreCase)
        {
            ["StkVendorRcvList"] = new[] { "StkVendorReceivingList", "RcvList" },
            ["StkVendorReceivingList"] = new[] { "StkVendorRcvList", "RcvList" },
            ["RcvList"] = new[] { "StkVendorRcvList", "StkVendorReceivingList" },

            ["StkVendorGateRcv"] = new[] { "StkVendorGateRcvd", "StkVendorGateReceiving", "PORcv" },
            ["StkVendorGateRcvd"] = new[] { "StkVendorGateRcv", "StkVendorGateReceiving", "PORcv" },
            ["PORcv"] = new[] { "StkVendorGateRcv", "StkVendorGateRcvd", "StkVendorGateReceiving" },

            ["StkRMLedger"] = new[] { "StkStockLedger" },
            ["StkStockLedger"] = new[] { "StkRMLedger" },

            ["StkSFOpenRcv"] = new[] { "StkSemiFinishOpenReceiving" },
            ["StkSemiFinishOpenReceiving"] = new[] { "StkSFOpenRcv" },

            ["StkFinishIssuance"] = new[] { "StkFinishStockIssuance" },
            ["StkFinishStockIssuance"] = new[] { "StkFinishIssuance" },

            ["StkFinishReceiving"] = new[] { "StkFinishStockReceiving" },
            ["StkFinishStockReceiving"] = new[] { "StkFinishReceiving" },

            ["PrdLotIssuance"] = new[] { "LotIssuance" },
            ["PrdReceiveLot"] = new[] { "ReceiveLot" },
            ["PrdReceivingList"] = new[] { "ReceivingList" },
            ["PrdMakerPO"] = new[] { "MakerPO" },
            ["PrdMakerPOList"] = new[] { "MakerPOList" },

            // Export Aliases (mapping Blazor Exp* IDs to legacy database MenuOptions OptionIDs and vice-versa)
            ["ExpOrderList"] = new[] { "OrderList", "OrderListCustomers", "OrderListStock" },
            ["OrderList"] = new[] { "ExpOrderList" },
            ["OrderListCustomers"] = new[] { "ExpOrderList" },
            ["OrderListStock"] = new[] { "ExpOrderList" },

            ["ExpOrderEntry"] = new[] { "NewOrder", "OrderList" },
            ["NewOrder"] = new[] { "ExpOrderEntry" },

            ["ExpCustomer"] = new[] { "FCustomers", "Customers", "CmpCustomerList", "OrderList", "NewOrder" },
            ["FCustomers"] = new[] { "ExpCustomer" },

            ["ExpQuotationList"] = new[] { "QuotationList" },
            ["QuotationList"] = new[] { "ExpQuotationList" },

            ["ExpAdvancePayment"] = new[] { "AdvancePayments" },
            ["AdvancePayments"] = new[] { "ExpAdvancePayment" },

            ["ExpOrderItemList"] = new[] { "OrderItemList" },
            ["OrderItemList"] = new[] { "ExpOrderItemList" },

            ["ExpCustomerItemBalances"] = new[] { "OrderItemBalances" },
            ["OrderItemBalances"] = new[] { "ExpCustomerItemBalances" },

            ["ExpArticlewiseShipped"] = new[] { "ArticleWiseStatus" },
            ["ArticleWiseStatus"] = new[] { "ExpArticlewiseShipped" },

            ["ExpProforma"] = new[] { "NewProforma", "PInvoiceList" },
            ["NewProforma"] = new[] { "ExpProforma" },

            ["ExpProformaList"] = new[] { "PInvoiceList", "NewProforma" },
            ["PInvoiceList"] = new[] { "ExpProformaList" },

            ["ExpCustomInvoice"] = new[] { "CustomInvoice", "NewCustomInvoice" },
            ["CustomInvoice"] = new[] { "ExpCustomInvoice" },

            ["ExpNewCustomInvoice"] = new[] { "NewCustomInvoice", "CustomInvoice" },
            ["NewCustomInvoice"] = new[] { "ExpNewCustomInvoice" },

            ["ExpCustomPaymentStatus"] = new[] { "CustPayStatus", "RecCustPay" },
            ["CustPayStatus"] = new[] { "ExpCustomPaymentStatus" },

            ["ExpReceiveCustomPayment"] = new[] { "RecCustPay", "CustPayStatus" },
            ["RecCustPay"] = new[] { "ExpReceiveCustomPayment" },

            ["ExpCommercialInvoice"] = new[] { "ComInvoice" },
            ["ComInvoice"] = new[] { "ExpCommercialInvoice" },

            ["ExpBankInvoice"] = new[] { "PrintInvoice" },
            ["PrintInvoice"] = new[] { "ExpBankInvoice" },

            ["ExpPackingList"] = new[] { "CustomLabels", "PrintInnerLabels", "ComPackingList", "CustomPackingList", "PrintLabels", "PrintPList", "NewPackingListM" },
            ["CustomLabels"] = new[] { "ExpPackingList" },
            ["PrintInnerLabels"] = new[] { "ExpPackingList" },
            ["ComPackingList"] = new[] { "ExpPackingList" },
            ["CustomPackingList"] = new[] { "ExpPackingList" },
            ["PrintLabels"] = new[] { "ExpPackingList" },
            ["PrintPList"] = new[] { "ExpPackingList" },
            ["NewPackingListM"] = new[] { "ExpPackingList" },

            ["ExpShippingInstructions"] = new[] { "CustomShipping" },
            ["CustomShipping"] = new[] { "ExpShippingInstructions" },

            ["ExpValuationForm"] = new[] { "PrintValuationForm" },
            ["PrintValuationForm"] = new[] { "ExpValuationForm" },

            ["ExpCommercialCovering"] = new[] { "ComCovering" },
            ["ComCovering"] = new[] { "ExpCommercialCovering" }
        };

        public bool HasOptionAccess(string optionId)
        {
            if (IsAdministrator) return true;
            if (string.IsNullOrWhiteSpace(optionId)) return true;

            if (_allowedOptionIds.Contains(optionId)) return true;

            if (OptionAliases.TryGetValue(optionId, out var aliases))
            {
                foreach (var alias in aliases)
                {
                    if (_allowedOptionIds.Contains(alias)) return true;
                }
            }

            return false;
        }

                private static readonly Dictionary<string, (string Module, string? OptionId)> RoutePermissions = new(StringComparer.OrdinalIgnoreCase)
        {
["accounts/accountgroups"] = ("Accounts", "AccExpenseGroups"),
            ["accounts/accountsledger"] = ("Accounts", "AccLedger"),
            ["accounts/bankbalancestatement"] = ("Accounts", "AccCashBankStatus"),
            ["accounts/banklist"] = ("Accounts", "BankList"),
            ["accounts/bankpaymentvoucher"] = ("Accounts", "AccPV"),
            ["accounts/bankreceiptvoucher"] = ("Accounts", "AccRV"),
            ["accounts/cashbookreport"] = ("Accounts", "AccCashBook"),
            ["accounts/cashpaymentvoucher"] = ("Accounts", "AccEV"),
            ["accounts/cashreceiptvoucher"] = ("Accounts", "AccRV"),
            ["accounts/changecategory"] = ("Accounts", "AccChangeAccHeads"),
            ["accounts/chartofaccounts"] = ("Accounts", "AccChartOfAccounts"),
            ["accounts/chqbookdetail"] = ("Accounts", "AccChqBookDetail"),
            ["accounts/custominvoiceauth"] = ("Accounts", "AccCustomInvoiceAuth"),
            ["accounts/journalvoucher"] = ("Accounts", "AccJV"),
            ["accounts/maker-loan-clearance"] = ("Accounts", "AccMakerLoanClearance"),
            ["accounts/maker-loan-transfer"] = ("Accounts", "AccMakerLoanTransfer"),
            ["accounts/maker-long-term-loan"] = ("Accounts", "AccMakerLoan"),
            ["accounts/maker-long-term-loan-ledger"] = ("Accounts", "AccMakerLoanLedger"),
            ["accounts/maker-short-term-loan"] = ("Accounts", "AccMakerLoan"),
            ["accounts/maker-short-term-loan-ledger"] = ("Accounts", "AccMakerLoanLedger"),
            ["accounts/payable"] = ("Accounts", "AccPayables"),
            ["accounts/receivable"] = ("Accounts", "AccReceivables"),
            ["accounts/reindex"] = ("Accounts", "AccRe-Index"),
            ["accounts/transactionregister"] = ("Accounts", "AccTransactionRegister"),
            ["accounts/trialbalance"] = ("Accounts", "AccTBSummary"),
            ["company/currency-exchange-rates"] = ("Company", "CmpExchangeRates"),
            ["company/customer-catalog"] = ("Company", "CmpCustomerCatalog"),
            ["company/item-groups"] = ("Company", "CmpItemGroups"),
            ["company/items"] = ("Company", "CmpItems"),
            ["company/misc-setup/additionalgroups"] = ("Company", "CmpItemGroups"),
            ["company/misc-setup/itemfinishedquality"] = ("Company", "CmpItemGroups"),
            ["company/misc-setup/itemtypes"] = ("Company", "CmpItemGroups"),
            ["company/new-item"] = ("Company", "CmpNewItem"),
            ["company/ports"] = ("Company", "CmpPorts"),
            ["company/steel-list"] = ("Company", "CmpSteelList"),
            ["company/stores-racks-bins"] = ("Company", "CmpStores"),
            ["dashboards/command-center"] = ("DashBoard", "DshCommandCenterGrid"),
            ["dashboards/command-center-analytics"] = ("DashBoard", "DshCommandCenter"),
            ["dashboards/executive"] = ("DashBoard", "DshExecutive"),
            ["dashboards/production-planning"] = ("DashBoard", "DshProductionPlanningGrid"),
            ["dashboards/production-planning-analytics"] = ("DashBoard", "DshProductionPlanning"),
            ["export/advance-payment-list"] = ("Export", "ExpAdvancePayment"),
            ["export/custom-payment-status"] = ("Export", "ExpCustomPaymentStatus"),
            ["export/customer"] = ("Export", "ExpCustomer"),
            ["export/customer-order-list"] = ("Export", "ExpOrderList"),
            ["export/customer-quotation-list"] = ("Export", "ExpQuotationList"),
            ["export/customers"] = ("Export", "ExpCustomer"),
            ["export/custominvoices/new-custominvoice"] = ("Export", "ExpNewCustomInvoice"),
            ["export/invoices/bank"] = ("Export", "ExpBankInvoice"),
            ["export/invoices/commercial"] = ("Export", "ExpCommercialInvoice"),
            ["export/invoices/commercial-covering"] = ("Export", "ExpCommercialCovering"),
            ["export/invoices/custom"] = ("Export", "ExpCustomInvoice"),
            ["export/invoices/packing-labels"] = ("Export", "ExpPackingList"),
            ["export/invoices/print-inner-labels"] = ("Export", "ExpPackingList"),
            ["export/invoices/print-valuation-form"] = ("Export", "ExpValuationForm"),
            ["export/invoices/shipping-instructions"] = ("Export", "ExpShippingInstructions"),
            ["export/new-customer-order"] = ("Export", "ExpOrderEntry"),
            ["export/order-item-list"] = ("Export", "ExpOrderItemList"),
            ["export/orders/articlewise-shipped-status"] = ("Export", "ExpArticlewiseShipped"),
            ["export/orders/customer-item-balances"] = ("Export", "ExpCustomerItemBalances"),
            ["export/proformas/new-proforma"] = ("Export", "ExpProforma"),
            ["export/proformas/proforma-list"] = ("Export", "ExpProformaList"),
            ["export/receive-custom-payment"] = ("Export", "ExpReceiveCustomPayment"),
            ["export/statistics"] = ("Export", null),
            ["export/statistics/articlewise-demand"] = ("Export", "StatArticleWiesOrders"),
            ["export/statistics/articlewise-sales"] = ("Export", "StatArticleWiesSales"),
            ["export/statistics/itemwise-export-performance"] = ("Export", "StatExportPerformanceReport"),
            ["export/statistics/total-demand"] = ("Export", "StatTotalDemand"),
            ["export/statistics/total-export"] = ("Export", "StatTotalExport"),
            ["intraoffice/health"] = ("Setup", "IntraOfficeHealth"),
            ["newrm"] = ("Stock", "StkNewRM"),
            ["newvendor"] = ("Stock", "StkNewVendor"),
            ["office"] = ("IntraOffice", "OfficeForms"),
            ["office/admin/email-settings"] = ("IntraOffice", "OfficeEmailSettings"),
            ["office/admin/minute-types"] = ("IntraOffice", "OfficeMinuteTypes"),
            ["office/ai-assistant"] = ("IntraOffice", "OfficeAI"),
            ["office/announcements"] = ("IntraOffice", "OfficeAnnouncements"),
            ["office/chat"] = ("IntraOffice", "OfficeChat"),
            ["office/directory"] = ("IntraOffice", "OfficeDirectory"),
            ["office/hub"] = ("IntraOffice", "OfficeHub"),
            ["office/leads"] = ("IntraOffice", "OfficeLeads"),
            ["office/meetings"] = ("IntraOffice", "OfficeMeetings"),
            ["office/messages"] = ("IntraOffice", "OfficeMessages"),
            ["office/minutes-list"] = ("IntraOffice", "OfficeMinutes"),
            ["office/reports"] = ("IntraOffice", "OfficeReports"),
            ["office/tasks"] = ("IntraOffice", "OfficeTasks"),
            ["office/templates"] = ("IntraOffice", "OfficeTemplates"),
            ["payroll/absent-sheet"] = ("Payroll", "PayAbsentSheet"),
            ["payroll/adjust-deduction-amount"] = ("Payroll", "PayDeductionAdjust"),
            ["payroll/advance-ledger"] = ("Payroll", "PayAdvanceRegister"),
            ["payroll/advance-long"] = ("Payroll", "PayAdvances"),
            ["payroll/advance-short"] = ("Payroll", "PayAdvances"),
            ["payroll/attendance-ledger"] = ("Payroll", "PayAttendance"),
            ["payroll/auto-attendance"] = ("Payroll", "PayAttendanceAuto"),
            ["payroll/clear-long-term-loan"] = ("Payroll", "PayLoan"),
            ["payroll/clear-short-term-loan"] = ("Payroll", "PayLoan"),
            ["payroll/daily-activity"] = ("Payroll", "PayDailyActivity"),
            ["payroll/deptlist"] = ("Payroll", "PayDepartment"),
            ["payroll/designations"] = ("Payroll", "PayDesignation"),
            ["payroll/emp-fine-ledger"] = ("Payroll", "PayFineRegister"),
            ["payroll/emp-to-emp-transfer"] = ("Payroll", "PayEmpTransfer"),
            ["payroll/employee-transfer"] = ("Payroll", "PayEmpTransfer"),
            ["payroll/employeelist"] = ("Payroll", "PayEmpProfile"),
            ["payroll/eobi"] = ("Payroll", "PayEOBI"),
            ["payroll/gate-pass"] = ("Payroll", "PayGatePass"),
            ["payroll/gratuity-calculation"] = ("Payroll", "PayGratuity"),
            ["payroll/hold-salary"] = ("Payroll", "PayHoldSalary"),
            ["payroll/holidays"] = ("Payroll", "PayHolidays"),
            ["payroll/leaves"] = ("Payroll", "PayLeaves"),
            ["payroll/manual-attendance"] = ("Payroll", "PayAttendanceManual"),
            ["payroll/monthly-attendance"] = ("Payroll", "PayAttendanceMonthly"),
            ["payroll/newemployee"] = ("Payroll", "PayNewEmp"),
            ["payroll/overtime-authorization"] = ("Payroll", "PayOvertimeAuth"),
            ["payroll/policies"] = ("Payroll", "PayPolicies"),
            ["payroll/post-fine"] = ("Payroll", "PayPostFine"),
            ["payroll/reports"] = ("Payroll", "PayReports"),
            ["payroll/salary-history"] = ("Payroll", "PaySalaryCalculation"),
            ["payroll/salary-sheet"] = ("Payroll", "PaySalarySheet"),
            ["payroll/settings"] = ("Payroll", "PaySettings"),
            ["payroll/short-term-sheet"] = ("Payroll", "PayAdvances"),
            ["payroll/social-security"] = ("Payroll", "PaySocialSecurity"),
            ["production/authorize-received"] = ("Production", "PrdAuthReceived"),
            ["production/create-dispatch-list"] = ("Production", "PrdCreateDispatchList"),
            ["production/dispatch-list"] = ("Production", "PrdDispatchList"),
            ["production/item-list"] = ("Production", "PrdItemList"),
            ["production/lot-issuance"] = ("Production", "PrdLotIssuance"),
            ["production/maker-billing"] = ("Production", "PrdMakerBilling"),
            ["production/maker-billing-list"] = ("Production", "PrdMakerBillingList"),
            ["production/maker-issuance-from-sf"] = ("Production", "PrdMakerIssuanceSF"),
            ["production/maker-item-assignment"] = ("Production", "PrdMakerItemAssign"),
            ["production/maker-list"] = ("Production", "PrdMakerList"),
            ["production/maker-po"] = ("Production", "PrdMakerPO"),
            ["production/maker-po-list"] = ("Production", "PrdMakerPOList"),
            ["production/new-maker"] = ("Production", "PrdNewMaker"),
            ["production/process-groups"] = ("Production", "PrdProcessGroups"),
            ["production/processes"] = ("Production", "PrdProcesses"),
            ["production/receive-against-po"] = ("Production", "PrdReceivePO"),
            ["production/receive-lot"] = ("Production", "PrdReceiveLot"),
            ["production/receiving-list"] = ("Production", "PrdReceivingList"),
            ["production/repair-types"] = ("Production", "PrdRepairTypes"),
            ["production/rework-issuance"] = ("Production", "PrdMakerRework"),
            ["production/statistics"] = ("Production", "PrdStatistics"),
            ["production/transfer-to-ready-finish-stock"] = ("Production", "PrdTransferReadyFinish"),
            ["production/wastage-types"] = ("Production", "PrdWastageTypes"),
            ["rmlist"] = ("Stock", "StkRMList"),
            ["setup/users"] = ("Setup", "SetupUsers"),
            ["stock/change-batch-lot"] = ("Stock", "StkChangeBatchLot"),
            ["stock/change-batch-no"] = ("Stock", "StkChangeBatchNo"),
            ["stock/finish-item-ledger"] = ("Stock", "StkFinishItemLedger"),
            ["stock/finish-stock-issuance"] = ("Stock", "StkFinishIssuance"),
            ["stock/finish-stock-movement"] = ("Stock", "StkFinishMovement"),
            ["stock/finish-stock-receiving"] = ("Stock", "StkFinishReceiving"),
            ["stock/finish-transactions"] = ("Stock", "StkFinishTransactions"),
            ["stock/material-placement"] = ("Stock", "StkMaterialPlacement"),
            ["stock/material-placement-list"] = ("Stock", "StkMaterialPlacementList"),
            ["stock/new-rm-po"] = ("Stock", "StkNewRMPO"),
            ["stock/rm-issuance"] = ("Stock", "StkRMIssuance"),
            ["stock/rm-issuance-list"] = ("Stock", "StkRMIssuanceList"),
            ["stock/rm-movement"] = ("Stock", "StkMaterialMovement"),
            ["stock/rmgroups"] = ("Stock", "StkRMGroups"),
            ["stock/rmpolist"] = ("Stock", "StkRMPOList"),
            ["stock/semi-finish-open-receiving"] = ("Stock", "StkSFOpenRcv"),
            ["stock/sf-movement"] = ("Stock", "StkSFMovement"),
            ["stock/sf-transactions"] = ("Stock", "StkSFTransactions"),
            ["stock/stock-ledger"] = ("Stock", "StkRMLedger"),
            ["stock/vend-gate-rcvd"] = ("Stock", "StkVendorGateRcv"),
            ["stock/vend-rcv-list"] = ("Stock", "StkVendorRcvList"),
            ["stock/vendor-billing"] = ("Stock", "StkVendorBilling"),
            ["stock/vendor-billing-list"] = ("Stock", "StkVendorBillingList"),
            ["stock/vendor-rm-assignment"] = ("Stock", "StkVendorRMAssign"),
            ["vendorlist"] = ("Stock", "StkVendorList"),
        };

        public bool IsRouteAuthorized(string relativePath)
        {
            if (IsAdministrator) return true;
            if (string.IsNullOrWhiteSpace(relativePath)) return true;

            // Strip query string and leading/trailing slashes
            var cleanPath = relativePath.Split('?')[0].Trim('/');
            if (string.IsNullOrWhiteSpace(cleanPath)) return true;

            // Favourites hub is accessible to all authenticated users
            if (string.Equals(cleanPath, "favourites", StringComparison.OrdinalIgnoreCase)) return true;

                        // 1. Check exact match in configured dictionary
            if (RoutePermissions.TryGetValue(cleanPath, out var perm))
            {
                if (!HasModuleAccess(perm.Module)) return false;
                if (!string.IsNullOrEmpty(perm.OptionId) && !HasOptionAccess(perm.OptionId)) return false;
                return true;
            }

            // 1.2 Check parameterized subroutes or trailing path segments
            foreach (var kvp in RoutePermissions)
            {
                if (cleanPath.StartsWith(kvp.Key + "/", StringComparison.OrdinalIgnoreCase))
                {
                    if (!HasModuleAccess(kvp.Value.Module)) return false;
                    if (!string.IsNullOrEmpty(kvp.Value.OptionId) && !HasOptionAccess(kvp.Value.OptionId)) return false;
                    return true;
                }
            }

            // 2. Prefix matching for top-level module areas
            var segments = cleanPath.Split('/');
            var root = segments[0].ToLowerInvariant();

            return root switch
            {
                "financial" or "accounts" => HasModuleAccess("Accounts"),
                "payroll" => HasModuleAccess("Payroll"),
                "stock" or "inventory" => HasModuleAccess("Stock"),
                "production" => HasModuleAccess("Production"),
                "export" => HasModuleAccess("Export"),
                "company" => HasModuleAccess("Company"),
                "dashboards" => HasModuleAccess("DashBoard"),
                "setup" => HasModuleAccess("Setup"),
                _ => true
            };
        }

        public Task<List<MenuOptionModel>> GetMenuOptionsByModuleAsync(string moduleName)
            => _permissionData.GetMenuOptionsByModuleAsync(moduleName);

        public Task<List<string>> GetDistinctModulesAsync()
            => _permissionData.GetDistinctModulesAsync();

        public Task<HashSet<string>> GetUserMenuOptionIdsAsync(int userId)
            => _permissionData.GetUserMenuOptionIdsAsync(userId);

        public Task<bool> SaveUserMenuOptionsAsync(int userId, string moduleName, IEnumerable<string> selectedOptionIds)
            => _permissionData.SaveUserMenuOptionsAsync(userId, moduleName, selectedOptionIds);

        public Task<bool> SaveAllUserMenuOptionsAsync(int userId, IEnumerable<string> allOptionIds)
            => _permissionData.SaveAllUserMenuOptionsAsync(userId, allOptionIds);

        public Task<bool> SyncMenuOptionsAsync(IEnumerable<MenuOptionModel> options)
            => _permissionData.SyncMenuOptionsAsync(options);

        public async Task<bool> ResetAllOptionsAsync()
        {
            // Replicates legacy VB6 AddAllOptions()
            try
            {
                var coreOptions = GetStandardMenuOptionsCatalog();
                await _permissionData.SyncMenuOptionsAsync(coreOptions);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error resetting all menu options");
                return false;
            }
        }

        public static List<MenuOptionModel> GetStandardMenuOptionsCatalog()
        {
            return new List<MenuOptionModel>
            {
                // Accounts / Financial
                new() { OptionID = "AccNewAccount", OptionName = "Define New Account", ModuleName = "Accounts" },
                new() { OptionID = "AccPayables", OptionName = "Accounts Payable", ModuleName = "Accounts" },
                new() { OptionID = "AccReceivables", OptionName = "Accounts Receivable", ModuleName = "Accounts" },
                new() { OptionID = "AccChartOfAccounts", OptionName = "Chart Of Accounts", ModuleName = "Accounts" },
                new() { OptionID = "AccRe-Index", OptionName = "Financial Re-Indexing", ModuleName = "Accounts" },
                new() { OptionID = "AccFinancialClosing", OptionName = "Financial Closing", ModuleName = "Accounts" },
                new() { OptionID = "AccExpenseGroups", OptionName = "Groups Accounts", ModuleName = "Accounts" },
                new() { OptionID = "AccChangeAccHeads", OptionName = "Change Account Heads", ModuleName = "Accounts" },
                new() { OptionID = "BankList", OptionName = "Bank List", ModuleName = "Accounts" },
                new() { OptionID = "AccChqBookDetail", OptionName = "Bank Accounts & Chq. Books Detail", ModuleName = "Accounts" },
                new() { OptionID = "AccRV", OptionName = "Receipt Voucher", ModuleName = "Accounts" },
                new() { OptionID = "AccPV", OptionName = "Bank Payment Voucher", ModuleName = "Accounts" },
                new() { OptionID = "AccJV", OptionName = "Journal Voucher", ModuleName = "Accounts" },
                new() { OptionID = "AccEV", OptionName = "Cash Payment Voucher", ModuleName = "Accounts" },
                new() { OptionID = "AccLedger", OptionName = "Detailed Account Ledger", ModuleName = "Accounts" },
                new() { OptionID = "AccCashBankStatus", OptionName = "Cash & Bank Status", ModuleName = "Accounts" },
                new() { OptionID = "AccIncomeStatement", OptionName = "Income Statement", ModuleName = "Accounts" },
                new() { OptionID = "AccBalanceSheet", OptionName = "Balance Sheet", ModuleName = "Accounts" },
                new() { OptionID = "AccTBSummary", OptionName = "Trial Balance (Summary)", ModuleName = "Accounts" },
                new() { OptionID = "AccTBDetail", OptionName = "Trial Balance (Detail)", ModuleName = "Accounts" },
                new() { OptionID = "AccCashBook", OptionName = "Cash Book Report", ModuleName = "Accounts" },

                // Production
                new() { OptionID = "PrdLotIssuance", OptionName = "Lot Issuance", ModuleName = "Production" },
                new() { OptionID = "PrdReceiveLot", OptionName = "Receive Lot", ModuleName = "Production" },
                new() { OptionID = "PrdReceivingList", OptionName = "Receiving List", ModuleName = "Production" },
                new() { OptionID = "PrdMakerPO", OptionName = "Maker PO", ModuleName = "Production" },
                new() { OptionID = "PrdMakerPOList", OptionName = "Maker PO List", ModuleName = "Production" },
                new() { OptionID = "PrdReWorkIssuance", OptionName = "Re-Work Issuance", ModuleName = "Production" },
                new() { OptionID = "PrdAuthorizeReceived", OptionName = "Authorize Received Lots", ModuleName = "Production" },
                new() { OptionID = "PrdMakerIssuanceFromSF", OptionName = "Maker Issuance From SF", ModuleName = "Production" },
                new() { OptionID = "PrdMakerItemAssignment", OptionName = "Maker Item Assignment", ModuleName = "Production" },
                new() { OptionID = "PrdNewMaker", OptionName = "New Maker Setup", ModuleName = "Production" },
                new() { OptionID = "PrdTransferToReadyFinishStock", OptionName = "Transfer to Ready Finish Stock", ModuleName = "Production" },
                new() { OptionID = "PrdProductionItemList", OptionName = "Production Item List", ModuleName = "Production" },

                // Stock
                new() { OptionID = "StkRMIssuance", OptionName = "RM Issuance", ModuleName = "Stock" },
                new() { OptionID = "StkRMIssuanceList", OptionName = "RM Issuance List", ModuleName = "Stock" },
                new() { OptionID = "StkRMPOList", OptionName = "RM PO List", ModuleName = "Stock" },
                new() { OptionID = "StkVendorGateRcvd", OptionName = "Vendor Gate Receiving", ModuleName = "Stock" },
                new() { OptionID = "StkVendorReceivingList", OptionName = "Vendor Receiving List", ModuleName = "Stock" },
                new() { OptionID = "StkMaterialPlacement", OptionName = "Material Placement", ModuleName = "Stock" },
                new() { OptionID = "StkMaterialPlacementList", OptionName = "Material Placement List", ModuleName = "Stock" },
                new() { OptionID = "StkRMMovement", OptionName = "Material Movement", ModuleName = "Stock" },
                new() { OptionID = "StkStockLedger", OptionName = "RM Stock Ledger", ModuleName = "Stock" },
                new() { OptionID = "StkSemiFinishOpenReceiving", OptionName = "Semi Finish Open Receiving", ModuleName = "Stock" },
                new() { OptionID = "StkSFMovement", OptionName = "Semi Finish Material Movement", ModuleName = "Stock" },
                new() { OptionID = "StkSFTransactions", OptionName = "Semi Finish Transactions", ModuleName = "Stock" },
                new() { OptionID = "StkFinishStockIssuance", OptionName = "Finish Stock Issuance", ModuleName = "Stock" },
                new() { OptionID = "StkFinishStockReceiving", OptionName = "Finish Stock Receiving", ModuleName = "Stock" },
                new() { OptionID = "StkFinishMovement", OptionName = "Finish Movement", ModuleName = "Stock" },
                new() { OptionID = "StkFinishItemLedger", OptionName = "Finish Item Ledger", ModuleName = "Stock" },
                new() { OptionID = "StkFinishTransactions", OptionName = "Finish Transactions", ModuleName = "Stock" },

                // Payroll / HR
                new() { OptionID = "PrlEmployeeList", OptionName = "Employee Directory & Profiles", ModuleName = "Payroll" },
                new() { OptionID = "PrlAttendanceLedger", OptionName = "Attendance Ledger", ModuleName = "Payroll" },
                new() { OptionID = "PrlMonthlyAttendance", OptionName = "Monthly Attendance Processing", ModuleName = "Payroll" },
                new() { OptionID = "PrlSalarySheet", OptionName = "Monthly Salary Sheet", ModuleName = "Payroll" },
                new() { OptionID = "PrlSalaryHistory", OptionName = "Salary History Ledger", ModuleName = "Payroll" },
                new() { OptionID = "PrlEmployeeLeave", OptionName = "Employee Leave Management", ModuleName = "Payroll" },
                new() { OptionID = "PrlAdvanceShort", OptionName = "Short Term Salary Advance", ModuleName = "Payroll" },
                new() { OptionID = "PrlClearShortTermLoan", OptionName = "Clear Short Term Loans", ModuleName = "Payroll" },
                new() { OptionID = "PrlPostFine", OptionName = "Employee Fines & Deductions", ModuleName = "Payroll" },
                new() { OptionID = "PrlEmpFineLedger", OptionName = "Employee Fine Ledger", ModuleName = "Payroll" },
                new() { OptionID = "PrlEmpToEmpTransfer", OptionName = "Employee Transfer", ModuleName = "Payroll" },
                new() { OptionID = "PrlHoldSalary", OptionName = "Hold / Unhold Salary", ModuleName = "Payroll" },

                // Company Catalog
                new() { OptionID = "CmpCompanyDetail", OptionName = "Company Profile", ModuleName = "Company" },
                new() { OptionID = "CmpCompanyCatalog", OptionName = "Company Catalog", ModuleName = "Company" },
                new() { OptionID = "CmpItemGroups", OptionName = "Item Groups", ModuleName = "Company" },
                new() { OptionID = "CmpCustomerList", OptionName = "Customer List", ModuleName = "Company" },
                new() { OptionID = "CmpNewCustomer", OptionName = "New Customer", ModuleName = "Company" },
                new() { OptionID = "CmpCustomerCatalog", OptionName = "Customer Catalog", ModuleName = "Company" },

                // IntraOffice CRM
                new() { OptionID = "OfficeHub", OptionName = "Collaboration Hub", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeTasks", OptionName = "Team Tasks", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeMinutes", OptionName = "Executive Minutes", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeMeetings", OptionName = "Meetings & Video Rooms", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeAnnouncements", OptionName = "Announcements Board", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeChat", OptionName = "Team Channels & Discussions", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeMessages", OptionName = "Direct Messages", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeLeads", OptionName = "Leads & CRM Pipeline", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeReports", OptionName = "CRM Reports", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeAiAssistant", OptionName = "AI Assistant", ModuleName = "IntraOffice" }
            };
        }
    }
}
