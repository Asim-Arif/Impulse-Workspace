namespace Impulse.Constants
{
    public static class ReportNames
    {
        // Root Reports
        public const string Voucher = "Voucher.rpt";
        public const string BOD_Meetings_Participants = "BOD_Meetings_Participants.rpt";
        public const string EditedVouchers = "EditedVouchers.rpt";
        public const string TrialBalance = "TrialBalance.rpt";
        public const string CashnBankStatus = "CashnBankStatus.rpt";
        public const string ChequeReceivingList = "ChequeReceivingList.rpt";
        public const string ImportantTasksList = "ImportantTasksList.rpt";
        public const string ImportantTasks = "ImportantTasks.rpt";
        public const string BillsReceived = "BillsReceived.rpt";
        public const string BillsReceivedList = "BillsReceivedList.rpt";
        public const string NTN_CNIC_Nos_List = "NTN_CNIC_Nos.rpt";
        public const string LetterDispatchRegister = "Letter_Dispatch_Register.rpt";
        public const string CostCenterReport = "Cost_Center_Report.rpt";
        public const string AccountsBreakup = "Accounts_Breakup.rpt";
        public const string PrintedChequeHistory = "PrintedChequeHistory.rpt";
        public const string VoucherList_TaxPayable = "VoucherList_TaxPayable.rpt";
        public const string BelowTaxLimitReport = "BelowTaxLimitReport.rpt";
        public const string TaxDeductionReport = "TaxDeductionReport.rpt";
        public const string TaxExemptedReport = "TaxExemptedReport.rpt";
        public const string TaxNotApplicableReport = "TaxNotApplicableReport.rpt";
        public const string TaxDetailReportHeadwise = "TaxDetailReportHeadwise.rpt";
        public const string AccountsAgingReport = "Accounts_Again_Report.rpt";
        public const string DepreciationSchedule = "DepreciationSchedule.rpt";
        public const string Accounts_Datewise = "Accounts_Datewise.rpt";

        // Notes To Accounts
        public const string NotesToAccounts = "BalanceSheet_Notes_SIAL.rpt";
        public const string BalanceSheet = "rptBalanceSheet_SIAL.rpt";
        public const string ProfitandLoss = "rptProfitandLoss_SIAL.rpt";
        public const string CashFlowSummary = "CashFlowSummary.rpt";
        public const string CashFlowSummary_WOZ = "CashFlowSummary_WOZ.rpt";
        public const string CashFlowSummary_Headwise = "CashFlowSummary_Headwise.rpt";
        public const string CashFlowStatement = "CashFlowStatment_SIAL.rpt";

        // Accounts Subfolder
        public static class Accounts
        {
            private const string Folder = "";
            public const string TrialBalance = Folder + "TrialBalance.rpt";
            public const string Ledger = Folder + "GeneralLedger.rpt";
            public const string Cheques_M1 = Folder + "Cheques_M1.rpt";
            public const string LedgerWithTitle = Folder + "Ledger_WithOtherHead.rpt";
            public const string Transaction_Report = Folder + "Transaction_Report.rpt";
            public const string PrintLedgerReport = Folder + "Ledger.rpt";
            public const string PrintAccounts = Folder + "Accounts.rpt";
            public const string AccountsWithOpening = Folder + "AccountsWithOpening.rpt";
        }

        // Payroll Subfolder
        public static class Payroll
        {
            private const string Folder = "Payroll/";
            public const string PaySlip = Folder + "rptPayslip.rpt";

            // Employee List Reports
            public const string EmpList                = Folder + "rptEmpList.rpt";
            public const string DeptList               = Folder + "rptDeptList.rpt";
            public const string EmpListWithOldID       = Folder + "rptEmpListWithOldID.rpt";
            public const string EmpListDesignationwise = Folder + "rptEmpListDesignationwise.rpt";
            public const string EmpListWithoutSal      = Folder + "rptEmpListWithoutSal.rpt";
            public const string EmpCards               = Folder + "rptEmpCards.rpt";
            public const string EmpDetail              = Folder + "rptEmpDetail.rpt";
            public const string EmpCardFields          = Folder + "rptEmpCardFields.rpt";
            public const string EmpListWCardID         = Folder + "rptEmpListWCardID.rpt";
            public const string EmpListPercentage      = Folder + "rptEmpListPercentage.rpt";
            public const string AppForm                = Folder + "rptAppForm.rpt";
            public const string EmpSalHistory          = Folder + "rptEmpSalHistory.rpt";
            public const string EmpSalaryLedger        = Folder + "rptEmpSalLedger.rpt";
            public const string TerminatedEmployees        = Folder + "TerminatedEmployees.rpt";
            public const string NewHiredEmployees          = Folder + "NewHiredEmployees.rpt";
            public const string EmployeesTargetsPerformance = Folder + "Employees_Targets_Performance.rpt";
            public const string AdvanceSalaryReportFilled   = Folder + "AdvanceSalaryReport_Filled.rpt";
            public const string AdvanceSalarySlip           = Folder + "AdvanceSalary_Slip.rpt";
            public const string EmpLedger                  = Folder + "rptEmpLedgerWLTOT.rpt";


            public const string DailyOverTime          = Folder + "DailyOverTime.rpt";

            // Biometric Ledger / Daily Attendance Reports
            public const string MonthlyAttendanceStatus = Folder + "MonthlyAttendanceStatus.rpt";
            public const string MissingAttendance       = Folder + "MissingAttendance.rpt";
            public const string LateAttReportDeptwise   = Folder + "LateAttReportDeptwise.rpt";
            public const string DailyAttendance        = Folder + "rptDailyAttendance.rpt";


            public const string AbsentLeaveSummary     = Folder + "rptAbsentLeaveSummary.rpt";
            public const string LeavesList             = Folder + "LeavesList.rpt";
            public const string LateComers             = Folder + "LateComers.rpt";
            public const string PresentEmpList         = Folder + "rptPresentEmpList.rpt";
            public const string EmpLeaveBalanceMonthly = Folder + "rptEmpLeaveBalanceMonthly.rpt";
            public const string Holidays               = Folder + "rptHolidays.rpt";

            // Advance Reports
            public const string LongTermAdvanceReceipt = Folder + "LongTermAdvanceReceipt.rpt";

            // Salary Sheet & Payslip Reports
            public const string SalarySheet             = Folder + "SalarySheet.rpt";
            public const string SalarySheetDetail       = Folder + "rptSalarySheet.rpt";
            public const string SalarySheetUnFinal      = Folder + "rptSalarySheetUnFinal.rpt";
            public const string SalarySheetSummary      = Folder + "rptSalarySummary.rpt";
            public const string SalarySheetInternal     = Folder + "rptSalarySheetInternal.rpt";
            public const string EmpListNegativeSalaries = Folder + "EmpListNegativeSalaries.rpt";
            public const string LateEmpList             = Folder + "rptLateEmpList.rpt";
            public const string RupeeStructure          = Folder + "rptRupeeStructure.rpt";
            public const string PaySlipRule9            = Folder + "rptPayslip_Rule9.rpt";
            public const string PaySlipMini             = Folder + "rptPayslip_Mini.rpt";
            public const string PaySlipFull             = Folder + "rptPayslip_Full.rpt";
            public const string SalaryPacket            = Folder + "rptSalaryPacket.rpt";
            public const string PaySlipInternal         = Folder + "rptPayslipInternal.rpt";
            public const string OverTimePaymentSheet    = Folder + "rptOverTimePaymentSheet.rpt";

            // Social Security Reports
            public const string SocialSecuritySheet     = Folder + "rptSocialSecurity.rpt";
            public const string SocialSecurityForm      = Folder + "rptSSForm.rpt";
            public const string NonSocialSecurity       = Folder + "rptNonSocialSecurity.rpt";

            // EOBI Reports
            public const string EOBISheet               = Folder + "rptOldAge.rpt";
            public const string NonEOBI                = Folder + "rptNonEOBI.rpt";

            // Designation Reports
            public const string DesignationsList        = Folder + "Designations.rpt";
            public const string DesignationDetail      = Folder + "Designation.rpt";
        }


        // Export Subfolder
        public static class Export
        {
            private const string Folder = ""; // Since legacy reports are usually flat in RptPath
            public const string PartShippedItemwise = Folder + "rptPartiallyShippedOrdersItemwise.rpt";
            public const string CombinedOrderItemwise = Folder + "CombinedOrderReportItemwise.rpt";
            public const string PartShippedOrders = Folder + "rptPartiallyShippedOrders.rpt";
            public const string ProductionExportCost = Folder + "OrderwiseProductionExportCost.rpt";
            public const string FOrderList = Folder + "rptFOrderList.rpt";
            public const string PTCQEL = Folder + "PTCQEL.rpt";
            public const string PTC = Folder + "PTC.rpt";
            public const string CustOrderStatus = Folder + "rptCustOrderStatus.rpt";
            public const string ProdOrderMakerwise = Folder + "rptProductionOrderMakerwise.rpt";
            public const string ProdOrderGroupWise = Folder + "rptProductionOrderWranglerGroupWise.rpt";
            public const string ProdOrderWrangler1 = Folder + "rptProductionOrderWrangler_1.rpt";
            public const string ProdOrderWrangler = Folder + "rptProductionOrderWrangler.rpt";
            public const string ProductionOrder = Folder + "rptProductionOrder.rpt";
            public const string ProdOrderBalanceQty = Folder + "rptProductionOrderBalanceQty.rpt";
            public const string ProdOrderMarkingPlan = Folder + "rptProductionOrderWranglerMarkingPlan.rpt";
            public const string ProdOrderWValue = Folder + "rptProductionOrderWValueWrangler.rpt";
            public const string ProdOrderRM = Folder + "ProductionOrder_RM.rpt";
            public const string ProdOrderPacking = Folder + "ProductionOrder_Packing.rpt";
            public const string ProdOrderStatus = Folder + "rptProductionOrder_Status.rpt";
            public const string ProdOrderDetails = Folder + "rptProductionOrder_Details.rpt";
            public const string ProdOrderBalance = Folder + "Production_Order_Balance.rpt";
            public const string PrintOrderSlip = Folder + "Print_Order_Slip.rpt";
            public const string PrintOrderPictures = Folder + "Print_Order_Pictures.rpt";
            public const string CustomerPO = Folder + "CustomerPO.rpt";
            public const string OrderItemList = Folder + "OrderItemList.rpt";
            public const string ProductionTimeLine = Folder + "ProductionTimeLineReport.rpt";
            public const string UrgentItemsForm = Folder + "UrgetItemsRequestForm.rpt";
            public const string SemiFinishStock = Folder + "Order_SemiFinishStock.rpt";
            public const string FinishStock = Folder + "Order_FinishStock.rpt";
            public const string ScoreBoard = Folder + "ScoreBoard.rpt";
            public const string OrdersLogSheet = Folder + "Orders_Log_Sheet.rpt";
            public const string BalanceOrdersCT = Folder + "Balance_Orders_CT.rpt";
            public const string OrderStatusLog = Folder + "OrderStatusLog.rpt";
            public const string OrderStatusLogEx = Folder + "OrderStatusLog_Ex.rpt";
            public const string OrderPurchasePlanningLog = Folder + "OrderPurchasePlanningLog.rpt";
            public const string UrgentRequest = Folder + "UrgentRequest.rpt";
        }

        // Stock Subfolder
        public static class Stock
        {
            private const string Folder = ""; // Flat in RptPath
            public const string StockLedger = Folder + "StockLedger.rpt";
            public const string SFTrayCard = Folder + "SF_Tray_Card.rpt";
            public const string SFLedger = Folder + "SF_Ledger.rpt";
            public const string SFOpeningSlip = Folder + "SF_Opening_Slip.rpt";
            public const string SFTransactions = Folder + "SFStockTransactions.rpt";
            public const string SFStockDetail = Folder + "Items_SFStock_Detail.rpt";
            public const string BinWiseSFStockStatus = Folder + "BinWiseSemiFinishStockStatus.rpt";
            public const string RackWiseSFStockStatus = Folder + "RackWiseSemiFinishStockStatus.rpt";
            public const string DailyStockList = Folder + "DailyStockList.rpt";
            public const string DailyStockListZero = Folder + "DailyStockList_Zero.rpt";
            public const string CustomerCatalogWithDetails = Folder + "CustomerCatalogWithDetails.rpt";
            public const string CustStockList = Folder + "CustStockList.rpt";
            public const string CustPriceList = Folder + "rptCustomerPriceList.rpt";
            public const string StockTransactions = Folder + "StockTransactions.rpt";
            public const string FinishStockDetail = Folder + "Items_FinishedStock_Detail.rpt";
            public const string BinWiseFinishStockStatus = Folder + "BinWiseStockStatusOfItems_Lot_Batch_Wise.rpt";
            public const string RackWiseFinishStockStatus = Folder + "RackWiseFinishStockStatus.rpt";
            public const string FinishStockReceivingSlip = Folder + "Finished_Stock_Receiving_Slip.rpt";
            public const string FinishStockIssuanceSlip = Folder + "Finished_Stock_Issuance_Slip.rpt";
            public const string FinishItemLedger = Folder + "ItemLedger.rpt";
            public const string FinishItemLedgerBinCard = Folder + "ItemLedger_BC.rpt";
        }

        // Company Subfolder
        public static class Company
        {
            private const string Folder = ""; // Flat in RptPath
            public const string CombCustStockList = Folder + "CombCustStockList.rpt";
            public const string CustomerCatalog = Folder + "CustomerCatalog.rpt";
            public const string CustomerCatalogWithDetails = Folder + "CustomerCatalogWithDetails.rpt";
            public const string CustStockList = Folder + "CustStockList.rpt";
            public const string CustPriceList = Folder + "rptCustomerPriceList.rpt";
        }
    }
}
