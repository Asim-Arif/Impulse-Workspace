namespace CIP.Constants
{
    public static class ReportNames
    {
        // Root Reports
        public const string Voucher = "PV.rpt";
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
        //public const string CapitalWorkInProgress = "CapitalWorkInProgress.rpt";

        /// ///////////// Notes To Accounts
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
            private const string Folder = "Reports_Accounts/";
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
            public const string PaySlip = Folder + "PaySlip.rpt";
        }
    }
}
