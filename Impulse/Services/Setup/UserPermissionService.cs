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
        public bool ShowCustomerOrderNo => IsAdministrator || (_currentUserModel != null && _currentUserModel.Show_Customer_Order_No);

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
                "setup" or "setups" => _currentUserModel.SetupMainLink,
                _ => true
            };
        }

        private static readonly Dictionary<string, string[]> OptionAliases = new(StringComparer.OrdinalIgnoreCase)
        {
            // ==========================================
            // Stock / Inventory Aliases
            // ==========================================
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

            ["StkFinishMovement"] = new[] { "FinishMovement" },
            ["FinishMovement"] = new[] { "StkFinishMovement" },

            ["StkFinishItemLedger"] = new[] { "FinishItemLedger" },
            ["FinishItemLedger"] = new[] { "StkFinishItemLedger" },

            ["StkFinishTransactions"] = new[] { "FinishTransactions" },
            ["FinishTransactions"] = new[] { "StkFinishTransactions" },

            ["StkSFMovement"] = new[] { "SemiFinishMaterialMovement", "SFMovement" },
            ["SemiFinishMaterialMovement"] = new[] { "StkSFMovement" },
            ["SFMovement"] = new[] { "StkSFMovement" },

            ["StkSFTransactions"] = new[] { "SemiFinishTransactions", "SFTransactions" },
            ["SemiFinishTransactions"] = new[] { "StkSFTransactions" },
            ["SFTransactions"] = new[] { "StkSFTransactions" },

            ["StkMaterialMovement"] = new[] { "StkRMMovement", "MaterialMovement" },
            ["StkRMMovement"] = new[] { "StkMaterialMovement", "MaterialMovement" },
            ["MaterialMovement"] = new[] { "StkMaterialMovement", "StkRMMovement" },

            ["StkRMIssuance"] = new[] { "RMIssuance" },
            ["RMIssuance"] = new[] { "StkRMIssuance" },

            ["StkRMIssuanceList"] = new[] { "RMIssuanceList" },
            ["RMIssuanceList"] = new[] { "StkRMIssuanceList" },

            ["StkRMPOList"] = new[] { "RMPOList", "StkNewRMPO" },
            ["RMPOList"] = new[] { "StkRMPOList" },
            ["StkNewRMPO"] = new[] { "StkRMPOList", "RMPOList" },

            ["StkRMList"] = new[] { "RMList", "RawMaterialsList" },
            ["RMList"] = new[] { "StkRMList" },

            ["StkRMGroups"] = new[] { "StkMaterialGroup", "RMGroups" },
            ["StkMaterialGroup"] = new[] { "StkRMGroups", "RMGroups" },
            ["RMGroups"] = new[] { "StkMaterialGroup", "StkRMGroups" },

            ["StkNewRM"] = new[] { "NewRM" },
            ["NewRM"] = new[] { "StkNewRM" },

            ["StkNewVendor"] = new[] { "NewVendor" },
            ["NewVendor"] = new[] { "StkNewVendor" },

            ["StkVendorList"] = new[] { "VendorList" },
            ["VendorList"] = new[] { "StkVendorList" },

            ["StkVendorRMAssign"] = new[] { "VendorRMAssign" },
            ["VendorRMAssign"] = new[] { "StkVendorRMAssign" },

            ["StkMaterialPlacement"] = new[] { "MaterialPlacement" },
            ["MaterialPlacement"] = new[] { "StkMaterialPlacement" },

            ["StkMaterialPlacementList"] = new[] { "MaterialPlacementList" },
            ["MaterialPlacementList"] = new[] { "StkMaterialPlacementList" },

            ["StkVendorBilling"] = new[] { "StkVenderBilling", "VendorBilling" },
            ["StkVenderBilling"] = new[] { "StkVendorBilling", "VendorBilling" },
            ["VendorBilling"] = new[] { "StkVenderBilling", "StkVendorBilling" },

            ["StkVendorBillingList"] = new[] { "StkVenderBillingList", "VendorBillingList" },
            ["StkVenderBillingList"] = new[] { "StkVendorBillingList", "VendorBillingList" },
            ["VendorBillingList"] = new[] { "StkVenderBillingList", "StkVendorBillingList" },

            ["StkChangeBatchLot"] = new[] { "ChangeBatchLot" },
            ["ChangeBatchLot"] = new[] { "StkChangeBatchLot" },

            ["StkChangeBatchNo"] = new[] { "ChangeBatchNo" },
            ["ChangeBatchNo"] = new[] { "StkChangeBatchNo" },

            // ==========================================
            // Production Aliases
            // ==========================================
            ["PrdLotIssuance"] = new[] { "LotIssuance" },
            ["LotIssuance"] = new[] { "PrdLotIssuance" },

            ["PrdReceiveLot"] = new[] { "ReceiveLot" },
            ["ReceiveLot"] = new[] { "PrdReceiveLot" },

            ["PrdReceivingList"] = new[] { "ReceivingList" },
            ["ReceivingList"] = new[] { "PrdReceivingList" },

            ["PrdMakerPO"] = new[] { "MakerPO" },
            ["MakerPO"] = new[] { "PrdMakerPO" },

            ["PrdMakerPOList"] = new[] { "MakerPOList" },
            ["MakerPOList"] = new[] { "PrdMakerPOList" },

            ["PrdMakerRework"] = new[] { "PrdReWorkIssuance", "ReWorkIssuance" },
            ["PrdReWorkIssuance"] = new[] { "PrdMakerRework", "ReWorkIssuance" },
            ["ReWorkIssuance"] = new[] { "PrdMakerRework", "PrdReWorkIssuance" },

            ["PrdAuthReceived"] = new[] { "PrdAuthorizeReceived", "AuthorizeReceived" },
            ["PrdAuthorizeReceived"] = new[] { "PrdAuthReceived", "AuthorizeReceived" },
            ["AuthorizeReceived"] = new[] { "PrdAuthReceived", "PrdAuthorizeReceived" },

            ["PrdMakerIssuanceSF"] = new[] { "PrdMakerIssuanceFromSF", "MakerIssuanceFromSF" },
            ["PrdMakerIssuanceFromSF"] = new[] { "PrdMakerIssuanceSF", "MakerIssuanceFromSF" },
            ["MakerIssuanceFromSF"] = new[] { "PrdMakerIssuanceSF", "PrdMakerIssuanceFromSF" },

            ["PrdMakerItemAssign"] = new[] { "PrdMakerItemAssignment", "MakerItemAssignment" },
            ["PrdMakerItemAssignment"] = new[] { "PrdMakerItemAssign", "MakerItemAssignment" },
            ["MakerItemAssignment"] = new[] { "PrdMakerItemAssign", "PrdMakerItemAssignment" },

            ["PrdNewMaker"] = new[] { "NewMaker" },
            ["NewMaker"] = new[] { "PrdNewMaker" },

            ["PrdMakerList"] = new[] { "MakerList", "AccMakerList" },
            ["AccMakerList"] = new[] { "PrdMakerList", "MakerList" },
            ["MakerList"] = new[] { "PrdMakerList", "AccMakerList" },

            ["PrdTransferReadyFinish"] = new[] { "PrdTransferToReadyFinishStock", "TransferToReadyFinishStock" },
            ["PrdTransferToReadyFinishStock"] = new[] { "PrdTransferReadyFinish", "TransferToReadyFinishStock" },
            ["TransferToReadyFinishStock"] = new[] { "PrdTransferReadyFinish", "PrdTransferToReadyFinishStock" },

            ["PrdItemList"] = new[] { "PrdProductionItemList", "ProductionItemList" },
            ["PrdProductionItemList"] = new[] { "PrdItemList", "ProductionItemList" },
            ["ProductionItemList"] = new[] { "PrdItemList", "PrdProductionItemList" },

            ["PrdReceivePO"] = new[] { "ReceivePO", "ReceiveAgainstPO", "PrdReceivingAgainstPO" },
            ["PrdReceivingAgainstPO"] = new[] { "PrdReceivePO", "ReceivePO", "ReceiveAgainstPO" },
            ["ReceivePO"] = new[] { "PrdReceivePO", "PrdReceivingAgainstPO" },
            ["ReceiveAgainstPO"] = new[] { "PrdReceivePO", "PrdReceivingAgainstPO" },

            ["PrdCreateDispatchList"] = new[] { "CreateDispatchList" },
            ["CreateDispatchList"] = new[] { "PrdCreateDispatchList" },

            ["PrdDispatchList"] = new[] { "DispatchList" },
            ["DispatchList"] = new[] { "PrdDispatchList" },

            ["PrdMakerBilling"] = new[] { "MakerBilling" },
            ["MakerBilling"] = new[] { "PrdMakerBilling" },

            ["PrdMakerBillingList"] = new[] { "MakerBillingList" },
            ["MakerBillingList"] = new[] { "PrdMakerBillingList" },

            ["PrdProcesses"] = new[] { "Processes", "ProcessesSetup" },
            ["Processes"] = new[] { "PrdProcesses" },

            ["PrdProcessGroups"] = new[] { "ProcessGroups" },
            ["ProcessGroups"] = new[] { "PrdProcessGroups" },

            ["PrdRepairTypes"] = new[] { "RepairTypes" },
            ["RepairTypes"] = new[] { "PrdRepairTypes" },

            ["PrdWastageTypes"] = new[] { "WastageTypes" },
            ["WastageTypes"] = new[] { "PrdWastageTypes" },

            ["PrdStatistics"] = new[] { "ProductionStatistics" },
            ["ProductionStatistics"] = new[] { "PrdStatistics" },

            // ==========================================
            // Export Aliases
            // ==========================================
            ["ExpOrderList"] = new[] { "OrderList", "OrderListCustomers", "OrderListStock" },
            ["OrderList"] = new[] { "ExpOrderList" },
            ["OrderListCustomers"] = new[] { "ExpOrderList" },
            ["OrderListStock"] = new[] { "ExpOrderList" },

            ["ExpOrderEntry"] = new[] { "NewOrder", "ExpNewOrder", "OrderList" },
            ["NewOrder"] = new[] { "ExpOrderEntry" },
            ["ExpNewOrder"] = new[] { "ExpOrderEntry" },

            ["ExpCustomer"] = new[] { "FCustomers", "Customers", "CmpCustomerList", "ExpForeignCusts", "OrderList", "NewOrder" },
            ["FCustomers"] = new[] { "ExpCustomer" },
            ["Customers"] = new[] { "ExpCustomer" },
            ["ExpForeignCusts"] = new[] { "ExpCustomer" },

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

            ["ExpProforma"] = new[] { "NewProforma", "ExpNewPInvoice", "PInvoiceList" },
            ["NewProforma"] = new[] { "ExpProforma" },
            ["ExpNewPInvoice"] = new[] { "ExpProforma" },

            ["ExpProformaList"] = new[] { "PInvoiceList", "ExpPInvoiceList", "NewProforma" },
            ["PInvoiceList"] = new[] { "ExpProformaList" },
            ["ExpPInvoiceList"] = new[] { "ExpProformaList" },

            ["ExpCustomInvoice"] = new[] { "CustomInvoice", "NewCustomInvoice" },
            ["CustomInvoice"] = new[] { "ExpCustomInvoice" },

            ["ExpNewCustomInvoice"] = new[] { "NewCustomInvoice", "CustomInvoice" },
            ["NewCustomInvoice"] = new[] { "ExpNewCustomInvoice" },

            ["ExpCustomPaymentStatus"] = new[] { "CustPayStatus", "RecCustPay" },
            ["CustPayStatus"] = new[] { "ExpCustomPaymentStatus" },

            ["ExpReceiveCustomPayment"] = new[] { "RecCustPay", "CustPayStatus" },
            ["RecCustPay"] = new[] { "ExpReceiveCustomPayment" },

            ["ExpCommercialInvoice"] = new[] { "ComInvoice", "ExpComInvoice" },
            ["ComInvoice"] = new[] { "ExpCommercialInvoice" },
            ["ExpComInvoice"] = new[] { "ExpCommercialInvoice" },

            ["ExpBankInvoice"] = new[] { "PrintInvoice", "ExpPrintInvoice" },
            ["PrintInvoice"] = new[] { "ExpBankInvoice" },
            ["ExpPrintInvoice"] = new[] { "ExpBankInvoice" },

            ["ExpPackingList"] = new[] { "CustomLabels", "PrintInnerLabels", "ComPackingList", "CustomPackingList", "PrintLabels", "PrintPList", "NewPackingListM", "ExpPackingLabels" },
            ["CustomLabels"] = new[] { "ExpPackingList" },
            ["PrintInnerLabels"] = new[] { "ExpPackingList" },
            ["ComPackingList"] = new[] { "ExpPackingList" },
            ["CustomPackingList"] = new[] { "ExpPackingList" },
            ["PrintLabels"] = new[] { "ExpPackingList" },
            ["PrintPList"] = new[] { "ExpPackingList" },
            ["NewPackingListM"] = new[] { "ExpPackingList" },
            ["ExpPackingLabels"] = new[] { "ExpPackingList" },

            ["ExpShippingInstructions"] = new[] { "CustomShipping", "ExpCustomShipping" },
            ["CustomShipping"] = new[] { "ExpShippingInstructions" },
            ["ExpCustomShipping"] = new[] { "ExpShippingInstructions" },

            ["ExpValuationForm"] = new[] { "PrintValuationForm", "ExpValuationForm" },
            ["PrintValuationForm"] = new[] { "ExpValuationForm" },

            ["ExpCommercialCovering"] = new[] { "ComCovering", "ExpComCovering" },
            ["ComCovering"] = new[] { "ExpCommercialCovering" },
            ["ExpComCovering"] = new[] { "ExpCommercialCovering" },

            ["StatTotalExport"] = new[] { "TotalExport" },
            ["TotalExport"] = new[] { "StatTotalExport" },

            ["StatTotalDemand"] = new[] { "TotalDemand" },
            ["TotalDemand"] = new[] { "StatTotalDemand" },

            ["StatArticleWiesSales"] = new[] { "ArticlewiseSales" },
            ["ArticlewiseSales"] = new[] { "StatArticleWiesSales" },

            ["StatArticleWiesOrders"] = new[] { "ArticlewiseDemand" },
            ["ArticlewiseDemand"] = new[] { "StatArticleWiesOrders" },

            ["StatExportPerformanceReport"] = new[] { "ItemwiseExportPerformance" },
            ["ItemwiseExportPerformance"] = new[] { "StatExportPerformanceReport" },

            // ==========================================
            // Payroll Aliases
            // ==========================================
            ["PayEmpProfile"] = new[] { "PrlEmpList", "EmployeeList", "PrlEmployeeList", "PrlTempEmpList", "PrlEmpListWithSal", "PrlEmpListWithoutSal", "PrintEmpList", "PrlPrintEmpDetails" },
            ["PrlEmpList"] = new[] { "PayEmpProfile" },
            ["EmployeeList"] = new[] { "PayEmpProfile" },
            ["PrlEmployeeList"] = new[] { "PayEmpProfile" },

            ["PayNewEmp"] = new[] { "PrlNewEmp", "NewEmployee" },
            ["PrlNewEmp"] = new[] { "PayNewEmp" },
            ["NewEmployee"] = new[] { "PayNewEmp" },

            ["PayDepartment"] = new[] { "PrlDeptList", "DeptList", "PrlNewDept", "NewDept" },
            ["PrlDeptList"] = new[] { "PayDepartment" },
            ["DeptList"] = new[] { "PayDepartment" },

            ["PaySalaryCalculation"] = new[] { "PrlEmpSalHistory", "PrlSalaryHistory", "SalaryLedger", "PrlEmpSalaryLedger" },
            ["PrlEmpSalHistory"] = new[] { "PaySalaryCalculation" },
            ["PrlSalaryHistory"] = new[] { "PaySalaryCalculation" },
            ["SalaryLedger"] = new[] { "PaySalaryCalculation" },

            ["PayAttendance"] = new[] { "PrlEmpLedger", "PrlAttendanceLedger", "AttendanceRegister", "EmployeeLedger", "PrlAttendanceStatus" },
            ["PrlEmpLedger"] = new[] { "PayAttendance" },
            ["PrlAttendanceLedger"] = new[] { "PayAttendance" },
            ["EmployeeLedger"] = new[] { "PayAttendance" },
            ["AttendanceRegister"] = new[] { "PayAttendance" },

            ["PayEmpTransfer"] = new[] { "PrlTransferEmp", "PrlEmpToEmpTransfer", "PrlLTEmpToEmpTransfer" },
            ["PrlTransferEmp"] = new[] { "PayEmpTransfer" },
            ["PrlEmpToEmpTransfer"] = new[] { "PayEmpTransfer" },

            ["PayAttendanceManual"] = new[] { "PrlTakeAttendanceEx", "PrlManualAttendance", "TakeAttendance", "PrlTakeAttendance" },
            ["PrlTakeAttendanceEx"] = new[] { "PayAttendanceManual" },
            ["TakeAttendance"] = new[] { "PayAttendanceManual", "PayAttendanceAuto", "PayAttendanceMonthly" },

            ["PayOvertimeAuth"] = new[] { "PrlOverTimeEntry", "PrlAuthorizeOvertime", "AddOverTime", "OverTime" },
            ["PrlOverTimeEntry"] = new[] { "PayOvertimeAuth" },
            ["AddOverTime"] = new[] { "PayOvertimeAuth" },

            ["PayAttendanceAuto"] = new[] { "PrlTakeAttendanceAuto", "TakeAttendance" },
            ["PrlTakeAttendanceAuto"] = new[] { "PayAttendanceAuto" },

            ["PayDailyActivity"] = new[] { "PrlDailyActivitySheet", "PrlDailyLeavesReport" },
            ["PrlDailyActivitySheet"] = new[] { "PayDailyActivity" },

            ["PayLeaves"] = new[] { "PrlLeaves", "PrlEmployeeLeave", "PrlEmpWiseLeaves", "PrlEmpOnLeaves" },
            ["PrlLeaves"] = new[] { "PayLeaves" },
            ["PrlEmployeeLeave"] = new[] { "PayLeaves" },

            ["PayAttendanceMonthly"] = new[] { "PrlEditAttendance", "PrlMonthlyAttendance", "TakeAttendance" },
            ["PrlEditAttendance"] = new[] { "PayAttendanceMonthly" },
            ["PrlMonthlyAttendance"] = new[] { "PayAttendanceMonthly" },

            ["PayGatePass"] = new[] { "PrlGatePassEntries", "GatePass" },
            ["PrlGatePassEntries"] = new[] { "PayGatePass" },

            ["PayAdvances"] = new[] { "PrlShortTermLoan", "PrlLongTermLoan", "ShortTermLoan", "LongTermLoan", "PrlAdvanceShort", "PrlAdvanceLong", "AdvancesLedger", "CalculateAdvSal" },
            ["PrlShortTermLoan"] = new[] { "PayAdvances", "PayLoan" },
            ["PrlLongTermLoan"] = new[] { "PayAdvances", "PayLoan" },
            ["ShortTermLoan"] = new[] { "PayAdvances", "PayLoan" },
            ["LongTermLoan"] = new[] { "PayAdvances", "PayLoan" },
            ["PrlAdvanceShort"] = new[] { "PayAdvances" },

            ["PayDeductionAdjust"] = new[] { "PrlAdjDeduction", "PrlPerformanceDeductionAmt" },
            ["PrlAdjDeduction"] = new[] { "PayDeductionAdjust" },

            ["PayAdvanceRegister"] = new[] { "PrlShortTermLoanLedger", "PrlLongTermLoanLedger", "AdvancesLedger", "PrlLoanBalance", "PrlLoanBalanceDateRange", "PrlEmpLoanBalance" },
            ["PrlShortTermLoanLedger"] = new[] { "PayAdvanceRegister" },
            ["PrlLongTermLoanLedger"] = new[] { "PayAdvanceRegister" },
            ["AdvancesLedger"] = new[] { "PayAdvanceRegister", "PayAdvances" },

            ["PayPostFine"] = new[] { "PrlEmpNewFine", "PrlPostFine" },
            ["PrlEmpNewFine"] = new[] { "PayPostFine" },
            ["PrlPostFine"] = new[] { "PayPostFine" },

            ["PayFineRegister"] = new[] { "PrlEmpFineLedger" },
            ["PrlEmpFineLedger"] = new[] { "PayFineRegister" },

            ["PaySalarySheet"] = new[] { "PrlSalarySheetExt", "PrlSalarySheet", "CalculateSalSheet", "SalSheet" },
            ["PrlSalarySheetExt"] = new[] { "PaySalarySheet" },
            ["PrlSalarySheet"] = new[] { "PaySalarySheet" },
            ["CalculateSalSheet"] = new[] { "PaySalarySheet" },
            ["SalSheet"] = new[] { "PaySalarySheet" },

            ["PayHoldSalary"] = new[] { "PrlHoldSalary" },
            ["PrlHoldSalary"] = new[] { "PayHoldSalary" },

            ["PayGratuity"] = new[] { "PrlGratuity" },
            ["PrlGratuity"] = new[] { "PayGratuity" },

            ["PayAbsentSheet"] = new[] { "PrlAbsentSheet" },
            ["PrlAbsentSheet"] = new[] { "PayAbsentSheet" },

            ["PaySocialSecurity"] = new[] { "PrlSocialSecurity", "SocialSecurity", "PrlContSS" },
            ["PrlSocialSecurity"] = new[] { "PaySocialSecurity" },
            ["SocialSecurity"] = new[] { "PaySocialSecurity" },

            ["PayEOBI"] = new[] { "PrlEOBI", "OldAge", "PrlContEOBI" },
            ["PrlEOBI"] = new[] { "PayEOBI" },
            ["OldAge"] = new[] { "PayEOBI" },

            ["PayDesignation"] = new[] { "PrlDesignationsList", "Designations" },
            ["PrlDesignationsList"] = new[] { "PayDesignation" },
            ["Designations"] = new[] { "PayDesignation" },

            ["PaySettings"] = new[] { "PrlDepartmentSettings", "DeptSettings", "PrlMachineSettings", "PrlBarcodeSettings", "Settings" },
            ["PrlDepartmentSettings"] = new[] { "PaySettings" },
            ["DeptSettings"] = new[] { "PaySettings" },

            ["PayPolicies"] = new[] { "PrlPayrollPolicies", "TaxRanges" },
            ["PrlPayrollPolicies"] = new[] { "PayPolicies" },
            ["TaxRanges"] = new[] { "PayPolicies" },

            ["PayHolidays"] = new[] { "PrlHolidays", "DefineHolidays" },
            ["PrlHolidays"] = new[] { "PayHolidays" },
            ["DefineHolidays"] = new[] { "PayHolidays" },

            ["PayReports"] = new[] { "PrlReports", "PrintEmpCards", "ApplicationForm", "EmpPaySlip", "PrlApplicationForm", "PrlEmpCards" },
            ["PrlReports"] = new[] { "PayReports" },

            ["PayLoan"] = new[] { "PrlClearLoan", "PrlClearShortTermLoan", "ShortTermLoan", "LongTermLoan" },
            ["PrlClearLoan"] = new[] { "PayLoan" },
            ["PrlClearShortTermLoan"] = new[] { "PayLoan" },

            // ==========================================
            // Company Aliases
            // ==========================================
            ["CmpItems"] = new[] { "CompanyCatalog", "CmpCompanyCatalog", "CompanyInfo", "CmpCompanyDetail", "CompanyImportCatalog" },
            ["CompanyCatalog"] = new[] { "CmpItems", "CmpNewItem", "CmpCompanyCatalog" },
            ["CmpCompanyCatalog"] = new[] { "CmpItems", "CmpNewItem", "CompanyCatalog" },
            ["CompanyInfo"] = new[] { "CmpItems", "CmpCompanyDetail" },
            ["CmpCompanyDetail"] = new[] { "CompanyInfo", "CmpItems" },

            ["CmpNewItem"] = new[] { "NewItem", "CompanyCatalog", "CmpCompanyCatalog" },
            ["NewItem"] = new[] { "CmpNewItem" },

            ["CmpItemGroups"] = new[] { "GroupList", "AdditionalGroupList", "ItemTypes", "ItemFinishedQuality" },
            ["GroupList"] = new[] { "CmpItemGroups" },
            ["AdditionalGroupList"] = new[] { "CmpItemGroups" },
            ["ItemTypes"] = new[] { "CmpItemGroups" },
            ["ItemFinishedQuality"] = new[] { "CmpItemGroups" },

            ["CmpPorts"] = new[] { "Ports", "ExpPorts" },
            ["Ports"] = new[] { "CmpPorts" },
            ["ExpPorts"] = new[] { "CmpPorts" },

            ["CmpStores"] = new[] { "AddStoresRacksBins", "Stores" },
            ["AddStoresRacksBins"] = new[] { "CmpStores" },
            ["Stores"] = new[] { "CmpStores" },

            ["CmpSteelList"] = new[] { "SteelList", "ExpSteelList" },
            ["SteelList"] = new[] { "CmpSteelList" },
            ["ExpSteelList"] = new[] { "CmpSteelList" },

            ["CmpExchangeRates"] = new[] { "ExchangeRates", "CmpExchangeRates" },
            ["ExchangeRates"] = new[] { "CmpExchangeRates" },

            ["CmpCustomerCatalog"] = new[] { "CustomerCatalog", "ExpCustCatalog" },
            ["CustomerCatalog"] = new[] { "CmpCustomerCatalog" },
            ["ExpCustCatalog"] = new[] { "CmpCustomerCatalog" },

            ["CmpCustomerList"] = new[] { "FCustomers", "Customers", "ExpCustomer" },
            ["CmpNewCustomer"] = new[] { "NewCustomer", "FCustomers", "ExpCustomer" },

            // ==========================================
            // Accounts / Financial Aliases
            // ==========================================
            ["AccChartOfAccounts"] = new[] { "ChartOfAccounts" },
            ["ChartOfAccounts"] = new[] { "AccChartOfAccounts" },

            ["AccPayables"] = new[] { "Payables", "AccountsPayable" },
            ["AccReceivables"] = new[] { "Receivables", "AccountsReceivable" },

            ["AccExpenseGroups"] = new[] { "ExpenseGroups", "GroupsAccounts" },
            ["AccChangeAccHeads"] = new[] { "ChangeAccHeads" },
            ["AccRe-Index"] = new[] { "ReIndex", "FinancialReIndexing" },

            ["BankList"] = new[] { "AccBanks" },
            ["AccBanks"] = new[] { "BankList" },

            ["AccChqBookDetail"] = new[] { "ChqBookDetail", "BankAccountsChqBooks" },

            ["AccEV"] = new[] { "CashPayment", "CashPaymentVoucher" },
            ["AccPV"] = new[] { "BankPayment", "BankPaymentVoucher" },
            ["AccRV"] = new[] { "ReceiptVoucher", "CashReceipt", "BankReceipt" },
            ["AccJV"] = new[] { "JournalVoucher" },

            ["AccMakerLoan"] = new[] { "MakerLoan", "MakerShortTermLoan", "MakerLongTermLoan" },
            ["AccMakerLoanLedger"] = new[] { "MakerLoanLedger" },
            ["AccMakerLoanClearance"] = new[] { "MakerLoanClearance" },
            ["AccMakerLoanTransfer"] = new[] { "MakerLoanTransfer" },
            ["AccCustomInvoiceAuth"] = new[] { "CustomInvoiceAuth" },

            ["AccLedger"] = new[] { "Ledger", "DetailedAccountLedger" },
            ["AccTransactionRegister"] = new[] { "TransactionRegister" },
            ["AccTBSummary"] = new[] { "TrialBalance", "TBSummary", "AccTBDetail" },
            ["AccCashBankStatus"] = new[] { "CashBankStatus", "BankBalanceStatement" },
            ["AccCashBook"] = new[] { "CashBook" },

            // ==========================================
            // Dashboards, Setup & IntraOffice Aliases
            // ==========================================
            ["SetupHub"] = new[] { "Setup", "Setups" },
            ["SetupUsers"] = new[] { "UserManagement", "PrlUserManager" },
            ["UserManagement"] = new[] { "SetupUsers" },
            ["PrlUserManager"] = new[] { "SetupUsers" },

            ["DshExecutive"] = new[] { "ExecutiveDashboard" },
            ["DshCommandCenter"] = new[] { "CommandCenter" },
            ["DshCommandCenterGrid"] = new[] { "CommandCenterGrid" },
            ["DshProductionPlanning"] = new[] { "ProductionPlanning" },
            ["DshProductionPlanningGrid"] = new[] { "ProductionPlanningGrid" },

            ["OfficeAI"] = new[] { "OfficeAiAssistant" },
            ["OfficeAiAssistant"] = new[] { "OfficeAI" },
            ["OfficeForms"] = new[] { "OfficeHub", "OfficeTasks", "OfficeChat", "OfficeMessages", "OfficeAnnouncements" },
            ["OfficeDirectory"] = new[] { "OfficeHub", "OfficeChat" },
            ["OfficeMinuteTypes"] = new[] { "OfficeMinutes" },
            ["OfficeTemplates"] = new[] { "OfficeReports", "OfficeLeads" },
            ["OfficeEmailSettings"] = new[] { "OfficeReports", "OfficeMessages" },
            ["IntraOfficeHealth"] = new[] { "Diagnostics", "Health", "SystemHealth" },
            ["ExpStatistics"] = new[] { "ExpOrderList", "ExpCustomer" },
            ["StkStockOrderAdjustment"] = new[] { "StkFinishIssuance", "StkFinishStockIssuance" }
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
            ["accounts/voucher-approval"] = ("Accounts", "AccVoucherApproval"),
            ["company/currency-exchange-rates"] = ("Company", "CmpExchangeRates"),
            ["company/customer-catalog"] = ("Company", "CustomerCatalog"),
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
            ["export/custominvoices/new-custominvoice"] = ("Export", "NewCustomInvoice"),
            ["export/invoices/bank"] = ("Export", "ExpBankInvoice"),
            ["export/invoices/commercial"] = ("Export", "ExpCommercialInvoice"),
            ["export/invoices/commercial-covering"] = ("Export", "ExpCommercialCovering"),
            ["export/invoices/custom"] = ("Export", "CustomInvoice"),
            ["export/invoices/packing-labels"] = ("Export", "ExpPackingList"),
            ["export/invoices/print-inner-labels"] = ("Export", "ExpPackingList"),
            ["export/invoices/print-valuation-form"] = ("Export", "PrintValuationForm"),
            ["export/invoices/shipping-instructions"] = ("Export", "ExpShippingInstructions"),
            ["export/new-customer-order"] = ("Export", "ExpOrderEntry"),
            ["export/order-item-list"] = ("Export", "OrderItemList"),
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
            ["intraoffice"] = ("IntraOffice", null),
            ["office"] = ("IntraOffice", null),
            ["office/dashboard"] = ("IntraOffice", null),
            ["office/admin/email-settings"] = ("IntraOffice", "OfficeEmailSettings"),
            ["office/admin/minute-types"] = ("IntraOffice", "OfficeMinuteTypes"),
            ["office/ai-assistant"] = ("IntraOffice", "OfficeAI"),
            ["office/announcements"] = ("IntraOffice", "OfficeAnnouncements"),
            ["office/channels"] = ("IntraOffice", "OfficeChat"),
            ["office/chat"] = ("IntraOffice", "OfficeChat"),
            ["office/customer-360"] = ("IntraOffice", "OfficeLeads"),
            ["office/directory"] = ("IntraOffice", "OfficeDirectory"),
            ["office/hub"] = ("IntraOffice", "OfficeHub"),
            ["office/leads"] = ("IntraOffice", "OfficeLeads"),
            ["office/meetings"] = ("IntraOffice", "OfficeMeetings"),
            ["office/messages"] = ("IntraOffice", "OfficeMessages"),
            ["office/minutes-approval"] = ("IntraOffice", "OfficeMinutes"),
            ["office/minutes-list"] = ("IntraOffice", "OfficeMinutes"),
            ["office/reports"] = ("IntraOffice", "OfficeReports"),
            ["office/tasks"] = ("IntraOffice", "OfficeTasks"),
            ["office/templates"] = ("IntraOffice", "OfficeTemplates"),
            ["chat"] = ("IntraOffice", "OfficeChat"),
            ["messages"] = ("IntraOffice", "OfficeMessages"),
            ["tasks"] = ("IntraOffice", "OfficeTasks"),
            ["payroll/absent-sheet"] = ("Payroll", "PrlAbsentSheet"),
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
            ["payroll/eobi"] = ("Payroll", "PrlEOBI"),
            ["payroll/gate-pass"] = ("Payroll", "PayGatePass"),
            ["payroll/gratuity-calculation"] = ("Payroll", "PayGratuity"),
            ["payroll/hold-salary"] = ("Payroll", "PayHoldSalary"),
            ["payroll/holidays"] = ("Payroll", "PayHolidays"),
            ["payroll/leaves"] = ("Payroll", "PayLeaves"),
            ["payroll/manual-attendance"] = ("Payroll", "PrlTakeAttendanceEx"),
            ["payroll/monthly-attendance"] = ("Payroll", "PayAttendanceMonthly"),
            ["payroll/newemployee"] = ("Payroll", "PayNewEmp"),
            ["payroll/overtime-authorization"] = ("Payroll", "PayOvertimeAuth"),
            ["payroll/policies"] = ("Payroll", "PrlPayrollPolicies"),
            ["payroll/post-fine"] = ("Payroll", "PayPostFine"),
            ["payroll/reports"] = ("Payroll", "PayReports"),
            ["payroll/salary-history"] = ("Payroll", "PaySalaryCalculation"),
            ["payroll/salary-sheet"] = ("Payroll", "PaySalarySheet"),
            ["payroll/settings"] = ("Payroll", "PaySettings"),
            ["payroll/short-term-sheet"] = ("Payroll", "PayAdvances"),
            ["payroll/social-security"] = ("Payroll", "PrlSocialSecurity"),
            ["production/authorize-received"] = ("Production", "PrdAuthReceived"),
            ["production/create-dispatch-list"] = ("Production", "PrdCreateDispatchList"),
            ["production/dispatch-list"] = ("Production", "PrdDispatchList"),
            ["production/item-list"] = ("Production", "PrdItemList"),
            ["production/issuance-deletion-approval"] = ("Production", "PrdMakerPOList"),
            ["production/lot-deletion-approval"] = ("Production", "PrdReceivingList"),
            ["production/lot-issuance"] = ("Production", "PrdLotIssuance"),
            ["production/maker-billing"] = ("Production", "PrdMakerBilling"),
            ["production/maker-billing-list"] = ("Production", "PrdMakerBillingList"),
            ["production/maker-issuance-from-sf"] = ("Production", "PrdMakerIssuanceSF"),
            ["production/maker-item-assignment"] = ("Production", "PrdMakerItemAssign"),
            ["production/maker-rate-approval"] = ("Production", "PrdMakerItemAssign"),
            ["production/maker-list"] = ("Production", "AccMakerList"),
            ["production/maker-po"] = ("Production", "PrdMakerPO"),
            ["production/maker-po-list"] = ("Production", "PrdMakerPOList"),
            ["production/new-maker"] = ("Production", "PrdNewMaker"),
            ["production/process-groups"] = ("Production", "PrdProcessGroups"),
            ["production/processes"] = ("Production", "PrdProcesses"),
            ["production/receive-against-po"] = ("Production", "PrdReceivingAgainstPO"),
            ["production/receive-lot"] = ("Production", "PrdReceiveLot"),
            ["production/receiving-list"] = ("Production", "PrdReceivingList"),
            ["production/repair-types"] = ("Production", "PrdRepairTypes"),
            ["production/rework-issuance"] = ("Production", "PrdMakerRework"),
            ["production/skip-process-approval"] = ("Production", "PrdLotIssuance"),
            ["production/statistics"] = ("Production", "PrdStatistics"),
            ["production/transfer-to-ready-finish-stock"] = ("Production", "PrdTransferReadyFinish"),
            ["production/wastage-types"] = ("Production", "PrdWastageTypes"),
            ["rmlist"] = ("Stock", "StkRMList"),
            ["setup"] = ("Setup", "SetupHub"),
            ["setups"] = ("Setup", "SetupHub"),
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
            ["stock/rmgroups"] = ("Stock", "StkMaterialGroup"),
            ["stock/rmpolist"] = ("Stock", "StkRMPOList"),
            ["stock/semi-finish-open-receiving"] = ("Stock", "StkSFOpenRcv"),
            ["stock/sf-movement"] = ("Stock", "StkSFMovement"),
            ["stock/sf-transactions"] = ("Stock", "StkSFTransactions"),
            ["stock/stock-ledger"] = ("Stock", "StkRMLedger"),
            ["stock/stock-order-adjustment"] = ("Stock", "StkFinishIssuance"),
            ["stock/vend-gate-rcvd"] = ("Stock", "StkVendorGateRcv"),
            ["stock/vend-rcv-list"] = ("Stock", "StkVendorRcvList"),
            ["stock/vendor-billing"] = ("Stock", "StkVenderBilling"),
            ["stock/vendor-billing-list"] = ("Stock", "StkVenderBillingList"),
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

            // 1.2 Check parameterized subroutes or trailing path segments (longest matching prefix first)
            var matchingKvp = RoutePermissions
                .Where(kvp => cleanPath.StartsWith(kvp.Key + "/", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(kvp => kvp.Key.Length)
                .FirstOrDefault();

            if (!string.IsNullOrEmpty(matchingKvp.Key))
            {
                if (!HasModuleAccess(matchingKvp.Value.Module)) return false;
                if (!string.IsNullOrEmpty(matchingKvp.Value.OptionId) && !HasOptionAccess(matchingKvp.Value.OptionId)) return false;
                return true;
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
                "office" or "intraoffice" => HasModuleAccess("IntraOffice"),
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

        public IEnumerable<string> GetOptionAliases(string optionId)
        {
            if (string.IsNullOrWhiteSpace(optionId)) return Enumerable.Empty<string>();
            return OptionAliases.TryGetValue(optionId, out var aliases) ? aliases : Enumerable.Empty<string>();
        }

        public bool IsOptionGranted(string optionId, ISet<string> userOptionIds)
        {
            if (string.IsNullOrWhiteSpace(optionId) || userOptionIds == null) return false;
            if (userOptionIds.Contains(optionId)) return true;
            if (OptionAliases.TryGetValue(optionId, out var aliases))
            {
                foreach (var alias in aliases)
                {
                    if (userOptionIds.Contains(alias)) return true;
                }
            }
            return false;
        }

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
                new() { OptionID = "AccMakerLoan", OptionName = "Maker Loans (Short & Long Term)", ModuleName = "Accounts" },
                new() { OptionID = "AccMakerLoanLedger", OptionName = "Maker Loan Ledgers", ModuleName = "Accounts" },
                new() { OptionID = "AccMakerLoanClearance", OptionName = "Maker Loan Clearance", ModuleName = "Accounts" },
                new() { OptionID = "AccMakerLoanTransfer", OptionName = "Maker Loan Transfer", ModuleName = "Accounts" },
                new() { OptionID = "AccCustomInvoiceAuth", OptionName = "Custom Invoice Authorization", ModuleName = "Accounts" },
                new() { OptionID = "AccVoucherApproval", OptionName = "Voucher Deletion Approvals", ModuleName = "Accounts" },
                new() { OptionID = "AccLedger", OptionName = "Detailed Account Ledger", ModuleName = "Accounts" },
                new() { OptionID = "AccTransactionRegister", OptionName = "Transaction Register", ModuleName = "Accounts" },
                new() { OptionID = "AccCashBankStatus", OptionName = "Cash & Bank Status", ModuleName = "Accounts" },
                new() { OptionID = "AccIncomeStatement", OptionName = "Income Statement", ModuleName = "Accounts" },
                new() { OptionID = "AccBalanceSheet", OptionName = "Balance Sheet", ModuleName = "Accounts" },
                new() { OptionID = "AccTBSummary", OptionName = "Trial Balance (Summary)", ModuleName = "Accounts" },
                new() { OptionID = "AccTBDetail", OptionName = "Trial Balance (Detail)", ModuleName = "Accounts" },
                new() { OptionID = "AccCashBook", OptionName = "Cash Book Report", ModuleName = "Accounts" },

                // Dashboards & Intelligence
                new() { OptionID = "DshExecutive", OptionName = "Executive Dashboards Suite", ModuleName = "DashBoard" },
                new() { OptionID = "DshCommandCenter", OptionName = "Command Center (Visual)", ModuleName = "DashBoard" },
                new() { OptionID = "DshCommandCenterGrid", OptionName = "Command Center (Grid)", ModuleName = "DashBoard" },
                new() { OptionID = "DshProductionPlanning", OptionName = "Production Planning (Visual)", ModuleName = "DashBoard" },
                new() { OptionID = "DshProductionPlanningGrid", OptionName = "Production Planning (Grid)", ModuleName = "DashBoard" },

                // Production
                new() { OptionID = "PrdOrderPlanning", OptionName = "PPC Order Planning", ModuleName = "Production" },
                new() { OptionID = "PrdMakerOrders", OptionName = "Maker POs (From PPC)", ModuleName = "Production" },
                new() { OptionID = "PrdOrderManagement", OptionName = "Order Management", ModuleName = "Production" },
                new() { OptionID = "PrdOrderLotsTracking", OptionName = "Order Lots & Hub Tracking", ModuleName = "Production" },
                new() { OptionID = "PrdEmployeePerformance", OptionName = "Worker Performance & Capacity", ModuleName = "Production" },
                new() { OptionID = "PrdCustOrderList", OptionName = "Customer Orders (Production)", ModuleName = "Production" },
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
                new() { OptionID = "AccMakerList", OptionName = "Maker List", ModuleName = "Production" },
                new() { OptionID = "PrdReceivingAgainstPO", OptionName = "Receive against PO", ModuleName = "Production" },
                new() { OptionID = "PrdTransferToReadyFinishStock", OptionName = "Transfer to Ready Finish Stock", ModuleName = "Production" },
                new() { OptionID = "PrdCreateDispatchList", OptionName = "Create Dispatch List", ModuleName = "Production" },
                new() { OptionID = "PrdDispatchList", OptionName = "Dispatch List", ModuleName = "Production" },
                new() { OptionID = "PrdMakerBilling", OptionName = "Maker Billing", ModuleName = "Production" },
                new() { OptionID = "PrdMakerBillingList", OptionName = "Maker Billing List", ModuleName = "Production" },
                new() { OptionID = "PrdProcesses", OptionName = "Processes Setup", ModuleName = "Production" },
                new() { OptionID = "PrdProcessGroups", OptionName = "Process Groups", ModuleName = "Production" },
                new() { OptionID = "PrdRepairTypes", OptionName = "Repair Types", ModuleName = "Production" },
                new() { OptionID = "PrdWastageTypes", OptionName = "Wastage Types", ModuleName = "Production" },
                new() { OptionID = "PrdStatistics", OptionName = "Production Statistics", ModuleName = "Production" },
                new() { OptionID = "PrdProductionItemList", OptionName = "Production Item List", ModuleName = "Production" },

                // Export
                new() { OptionID = "ExpCustomer", OptionName = "Foreign Customers", ModuleName = "Export" },
                new() { OptionID = "ExpOrderEntry", OptionName = "New Customer Order", ModuleName = "Export" },
                new() { OptionID = "ExpOrderList", OptionName = "Customer Order List", ModuleName = "Export" },
                new() { OptionID = "ExpQuotationList", OptionName = "Customer Quotation List", ModuleName = "Export" },
                new() { OptionID = "ExpAdvancePayment", OptionName = "Advance Payment List", ModuleName = "Export" },
                new() { OptionID = "OrderItemList", OptionName = "Order Item List", ModuleName = "Export" },
                new() { OptionID = "ExpCustomerItemBalances", OptionName = "Customer Item Balances", ModuleName = "Export" },
                new() { OptionID = "ExpArticlewiseShipped", OptionName = "Articlewise Shipped Status", ModuleName = "Export" },
                new() { OptionID = "ExpProforma", OptionName = "New Proforma Invoice", ModuleName = "Export" },
                new() { OptionID = "ExpProformaList", OptionName = "Proforma List", ModuleName = "Export" },
                new() { OptionID = "CustomInvoice", OptionName = "Custom Invoice List", ModuleName = "Export" },
                new() { OptionID = "NewCustomInvoice", OptionName = "New Custom Invoice", ModuleName = "Export" },
                new() { OptionID = "ExpCustomPaymentStatus", OptionName = "Custom Payment Status", ModuleName = "Export" },
                new() { OptionID = "ExpReceiveCustomPayment", OptionName = "Receive Custom Payment", ModuleName = "Export" },
                new() { OptionID = "ExpCommercialInvoice", OptionName = "Commercial Invoice List", ModuleName = "Export" },
                new() { OptionID = "ExpBankInvoice", OptionName = "Bank Invoice List", ModuleName = "Export" },
                new() { OptionID = "ExpPackingList", OptionName = "Packing Outer & Inner Labels", ModuleName = "Export" },
                new() { OptionID = "ExpShippingInstructions", OptionName = "Shipping Instructions", ModuleName = "Export" },
                new() { OptionID = "PrintValuationForm", OptionName = "Print Valuation Form", ModuleName = "Export" },
                new() { OptionID = "ExpCommercialCovering", OptionName = "Commercial Covering Letter", ModuleName = "Export" },
                new() { OptionID = "ExpStatistics", OptionName = "Export Statistics Hub", ModuleName = "Export" },

                // Stock
                new() { OptionID = "StkMaterialGroup", OptionName = "Raw Material Groups", ModuleName = "Stock" },
                new() { OptionID = "StkNewRM", OptionName = "New Raw Material", ModuleName = "Stock" },
                new() { OptionID = "StkRMList", OptionName = "Raw Materials List", ModuleName = "Stock" },
                new() { OptionID = "StkRMPOList", OptionName = "RM PO List", ModuleName = "Stock" },
                new() { OptionID = "StkNewRMPO", OptionName = "New RM PO", ModuleName = "Stock" },
                new() { OptionID = "StkNewVendor", OptionName = "New Vendor", ModuleName = "Stock" },
                new() { OptionID = "StkVendorList", OptionName = "Vendor List", ModuleName = "Stock" },
                new() { OptionID = "StkVendorRMAssign", OptionName = "Vendor RM Assignment", ModuleName = "Stock" },
                new() { OptionID = "StkVendorGateRcvd", OptionName = "Vendor Gate Receiving", ModuleName = "Stock" },
                new() { OptionID = "StkVendorReceivingList", OptionName = "Vendor Receiving List", ModuleName = "Stock" },
                new() { OptionID = "StkMaterialPlacement", OptionName = "Material Placement", ModuleName = "Stock" },
                new() { OptionID = "StkMaterialPlacementList", OptionName = "Material Placement List", ModuleName = "Stock" },
                new() { OptionID = "StkRMMovement", OptionName = "Material Movement", ModuleName = "Stock" },
                new() { OptionID = "StkChangeBatchLot", OptionName = "Change Batch / Lot", ModuleName = "Stock" },
                new() { OptionID = "StkChangeBatchNo", OptionName = "Change Batch No", ModuleName = "Stock" },
                new() { OptionID = "StkStockLedger", OptionName = "RM Stock Ledger", ModuleName = "Stock" },
                new() { OptionID = "StkRMIssuance", OptionName = "RM Issuance", ModuleName = "Stock" },
                new() { OptionID = "StkRMIssuanceList", OptionName = "RM Issuance List", ModuleName = "Stock" },
                new() { OptionID = "StkSemiFinishOpenReceiving", OptionName = "Semi Finish Open Receiving", ModuleName = "Stock" },
                new() { OptionID = "StkSFMovement", OptionName = "Semi Finish Material Movement", ModuleName = "Stock" },
                new() { OptionID = "StkSFTransactions", OptionName = "Semi Finish Transactions", ModuleName = "Stock" },
                new() { OptionID = "StkFinishStockIssuance", OptionName = "Finish Stock Issuance", ModuleName = "Stock" },
                new() { OptionID = "StkStockOrderAdjustment", OptionName = "Stock Order Adjustment (PPC)", ModuleName = "Stock" },
                new() { OptionID = "StkFinishStockReceiving", OptionName = "Finish Stock Receiving", ModuleName = "Stock" },
                new() { OptionID = "StkFinishMovement", OptionName = "Finish Movement", ModuleName = "Stock" },
                new() { OptionID = "StkFinishItemLedger", OptionName = "Finish Item Ledger", ModuleName = "Stock" },
                new() { OptionID = "StkFinishTransactions", OptionName = "Finish Transactions", ModuleName = "Stock" },
                new() { OptionID = "StkVenderBilling", OptionName = "Vendor Billing", ModuleName = "Stock" },
                new() { OptionID = "StkVenderBillingList", OptionName = "Vendor Billing List", ModuleName = "Stock" },

                // Payroll / HR
                new() { OptionID = "PayNewEmp", OptionName = "New Employee Profile", ModuleName = "Payroll" },
                new() { OptionID = "PrlEmployeeList", OptionName = "Employee Directory & Profiles", ModuleName = "Payroll" },
                new() { OptionID = "PayDepartment", OptionName = "Departments List", ModuleName = "Payroll" },
                new() { OptionID = "PrlAttendanceLedger", OptionName = "Attendance Ledger", ModuleName = "Payroll" },
                new() { OptionID = "PrlTakeAttendanceEx", OptionName = "Manual Attendance", ModuleName = "Payroll" },
                new() { OptionID = "PayOvertimeAuth", OptionName = "Overtime Authorization", ModuleName = "Payroll" },
                new() { OptionID = "PayAttendanceAuto", OptionName = "Auto Attendance", ModuleName = "Payroll" },
                new() { OptionID = "PayDailyActivity", OptionName = "Daily Activity", ModuleName = "Payroll" },
                new() { OptionID = "PrlMonthlyAttendance", OptionName = "Monthly Attendance Processing", ModuleName = "Payroll" },
                new() { OptionID = "PrlSalarySheet", OptionName = "Monthly Salary Sheet", ModuleName = "Payroll" },
                new() { OptionID = "PrlSalaryHistory", OptionName = "Salary History Ledger", ModuleName = "Payroll" },
                new() { OptionID = "PrlEmployeeLeave", OptionName = "Employee Leave Management", ModuleName = "Payroll" },
                new() { OptionID = "PayGatePass", OptionName = "Gate Pass Entry", ModuleName = "Payroll" },
                new() { OptionID = "PrlAdvanceShort", OptionName = "Short Term Salary Advance", ModuleName = "Payroll" },
                new() { OptionID = "PayAdvances", OptionName = "Long Term Salary Advance", ModuleName = "Payroll" },
                new() { OptionID = "PrlClearShortTermLoan", OptionName = "Clear Short Term Loans", ModuleName = "Payroll" },
                new() { OptionID = "PayLoan", OptionName = "Clear Long Term Loans", ModuleName = "Payroll" },
                new() { OptionID = "PrlPostFine", OptionName = "Employee Fines & Deductions", ModuleName = "Payroll" },
                new() { OptionID = "PrlEmpFineLedger", OptionName = "Employee Fine Ledger", ModuleName = "Payroll" },
                new() { OptionID = "PrlEmpToEmpTransfer", OptionName = "Employee Transfer", ModuleName = "Payroll" },
                new() { OptionID = "PayDeductionAdjust", OptionName = "Adjust Deduction Amount", ModuleName = "Payroll" },
                new() { OptionID = "PayAdvanceRegister", OptionName = "Loan & Advance Register", ModuleName = "Payroll" },
                new() { OptionID = "PrlHoldSalary", OptionName = "Hold / Unhold Salary", ModuleName = "Payroll" },
                new() { OptionID = "PayGratuity", OptionName = "Gratuity Calculation", ModuleName = "Payroll" },
                new() { OptionID = "PrlAbsentSheet", OptionName = "Absent Sheet", ModuleName = "Payroll" },
                new() { OptionID = "PrlSocialSecurity", OptionName = "Social Security Sheet", ModuleName = "Payroll" },
                new() { OptionID = "PrlEOBI", OptionName = "EOBI Sheet", ModuleName = "Payroll" },
                new() { OptionID = "PayDesignation", OptionName = "Designations Setup", ModuleName = "Payroll" },
                new() { OptionID = "PaySettings", OptionName = "Payroll Settings", ModuleName = "Payroll" },
                new() { OptionID = "PrlPayrollPolicies", OptionName = "Payroll Policies", ModuleName = "Payroll" },
                new() { OptionID = "PayHolidays", OptionName = "Holidays Setup", ModuleName = "Payroll" },
                new() { OptionID = "PayReports", OptionName = "Payroll Reports Hub", ModuleName = "Payroll" },

                // Company
                new() { OptionID = "CmpItems", OptionName = "Items List", ModuleName = "Company" },
                new() { OptionID = "CmpNewItem", OptionName = "New Item Setup", ModuleName = "Company" },
                new() { OptionID = "CmpCompanyDetail", OptionName = "Company Profile", ModuleName = "Company" },
                new() { OptionID = "CompanyCatalog", OptionName = "Company Catalog", ModuleName = "Company" },
                new() { OptionID = "CmpItemGroups", OptionName = "Item Groups & Classifications", ModuleName = "Company" },
                new() { OptionID = "CmpPorts", OptionName = "Ports & Locations", ModuleName = "Company" },
                new() { OptionID = "CmpStores", OptionName = "Stores, Racks & Bins", ModuleName = "Company" },
                new() { OptionID = "CmpSteelList", OptionName = "Steel List", ModuleName = "Company" },
                new() { OptionID = "CmpExchangeRates", OptionName = "Currency Exchange Rates", ModuleName = "Company" },
                new() { OptionID = "CmpCustomerList", OptionName = "Customer List", ModuleName = "Company" },
                new() { OptionID = "CmpNewCustomer", OptionName = "New Customer", ModuleName = "Company" },
                new() { OptionID = "CustomerCatalog", OptionName = "Customer Catalog", ModuleName = "Company" },

                // Setups & System Administration
                new() { OptionID = "SetupHub", OptionName = "Setups Hub", ModuleName = "Setup" },
                new() { OptionID = "SetupUsers", OptionName = "User Management", ModuleName = "Setup" },
                new() { OptionID = "OfficeMinuteTypes", OptionName = "Minute Types", ModuleName = "Setup" },
                new() { OptionID = "OfficeEmailSettings", OptionName = "Email & SMTP Settings", ModuleName = "Setup" },
                new() { OptionID = "IntraOfficeHealth", OptionName = "System Diagnostics", ModuleName = "Setup" },

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
                new() { OptionID = "OfficeAiAssistant", OptionName = "AI Assistant", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeDirectory", OptionName = "Employee Directory", ModuleName = "IntraOffice" },
                new() { OptionID = "OfficeTemplates", OptionName = "Email Templates", ModuleName = "IntraOffice" }
            };
        }
    }
}
