using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;
using CIP.Pages.Accounts;
using System.Security.Cryptography.X509Certificates;

namespace CIP.Services
{
    public interface IFinancialStatementService
    {
        Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "");
        Task<List<MainHeadsModel>> GetMainHeadsList();
        Task<List<SubHeadsModel>> GetSubHeadList();
        Task<List<SubHeadsModel>> GetMainHeads();
        Task<List<SubHeadsListModel>> GetSubHeads(long MainGroup_refID=0);

        Task<List<SubHeadAccountDetailModel>> GetAccountGroup_HeadsDetail(long AccGroup_RefID);
        Task<List<GroupsAccountsModel>> GetGroupsAccountsList();
        Task<List<AccountsHeadManualEntryModel>> GetAccountsHeadManualEntries();
        Task<List<AccountsHeadManualEntry_SC_Model>> GetAccountsHeadManualEntries_SC();
        Task<List<RemunerationandOtherBenefitsModel>> GetRemunerationandOtherBenefitsEntries();
        Task<List<ManualEntryForNotesModel>> GetManualEntryForNotesList();
        Task<List<SubHeadsManualEntryModel>> GetSubHeadsManualEntriesList();
        Task<List<Vouchers_Adjustment_Transaction_Model>> GetTransactionData_Adjustment(DateTime DTFrom, DateTime DTTo);
        Task<string> GetNextVchrNo_Adjustment(DateTime OnDate, string VchrType);
        Task<List<MainHeadsListModel>> GetAllHeads();
        Task<List<ProfitandLossHeadsDetailToLoadModel>> GetProfitandLossHeadsList();
        Task<List<BalanceSheetHeadsDetailToLoadModel>> GetBalanceSheetHeadsList();
        Task<List<FinancialStatementHeadsModel>> GetFinancialStatementHeadsList();
        Task<List<PrintAttachmentsModel>> GetNotesAttachmentList();
        Task<List<CashFlowMainHeadsModel>> GetCashFlowMainHeadsList();
        Task<List<CashFlowType_DD_Model>> GetCashFlowType_DD();
        Task<List<CashFlowMain_DD_Model>> GetCashFlowMainHeads_DD();
        Task<List<CashFlowReportHeadsDetailToLoadModel>> GetCashFlowReportHeadsList();
        Task<List<CashFlowHeadsListModel>> GetCashFlowHeadsList_DD();


        Task<bool> SaveMainHead(MainHeadsModel mainhead);
        Task<bool> SaveSubHead(SubHeadsModel subhead);
        Task<bool> SaveSubHeadDetail(SubHeadAccountDetailModel subhead);
        Task<bool> SaveGroupsAccounts(GroupsAccountsModel groupsaccounts);
        Task<bool> SaveAccountsHeadManualEntries(AccountsHeadManualEntryModel groupsaccounts);
        Task<bool> SaveAccountsHeadManualEntries_SC(AccountsHeadManualEntry_SC_Model groupsaccounts);
        Task<bool> SaveRemunerationandOtherBenefitsEntries(RemunerationandOtherBenefitsModel groupsaccounts);
        Task<bool> SaveManualEntryForNotes(ManualEntryForNotesModel groupsaccounts);
        Task<bool> SaveSubHeadsManualEntries(SubHeadsManualEntryModel groupsaccounts);
        Task<string> SaveAdjustmentVoucher(Vouchers_Main_Model voucheradjustmentmodel);
        Task<string> SaveProfitandLossHeads(ProfitandLossHeads_Main_Model profitandlossmain);
        Task<string> SaveBalanceSheetHeads(BalanceSheet_Main_Model balancesheet);
        Task<string> SaveFinancialStatementHeads(string username, string machinename);
        Task<int> SaveNotesAttachments(PrintAttachmentsModel groupsaccounts);
        Task<bool> SaveCashFlowMainHeads(CashFlowMainHeadsModel groupsaccounts);
        Task<string> SaveCashFlowReportHeads(CashFlowReportHeads_Main_Model cashflowmain);


        Task<bool> EditMainHead(MainHeadsModel mainhead,long entryid);
        Task<bool> EditSubHead(SubHeadsModel subhead, long entryid, int itype);
        Task<bool> EditSubHeadDetailEntry(SubHeadAccountDetailModel subhead, long entryid,int itype);
        Task<bool> EditGroupsAccounts(GroupsAccountsModel mainhead, long entryid);
        Task<bool> EditAccountsHeadManualEntries(AccountsHeadManualEntryModel mainhead, long entryid);
        Task<bool> EditAccountsHeadManualEntries_SC(AccountsHeadManualEntry_SC_Model mainhead, long entryid);
        Task<bool> EditRemunerationandOtherBenefitsEntries(RemunerationandOtherBenefitsModel mainhead, long entryid);
        Task<bool> EditManualEntryForNotes(ManualEntryForNotesModel mainhead, long entryid);
        Task<bool> EditSubHeadsManualEntries(SubHeadsManualEntryModel mainhead, long entryid);
        Task<bool> EditProfitandLossHeads(ProfitandLossHeads_Main_Model mainhead, long entryid);
        Task<bool> EditBalanceSheetHeads(BalanceSheet_Main_Model balancesheet, long entryid);
        Task<bool> EditFinancialStatementHeads(FinancialStatementHeadsModel mainhead, int itype, long entryid);
        Task<bool> EditCashFlowMainHeads(CashFlowMainHeadsModel mainhead, int itype, long entryid);
        Task<bool> EditCashFlowReportHeads(CashFlowReportHeads_Main_Model mainhead, long entryid);


        Task<bool> DeleteMainHead(long entryid);
        Task<bool> DeleteSubHead(long entryid);
        Task<bool> DeleteSubHeadEntry(long entryid);
        Task<bool> DeleteGroupsAccounts(long entryid);
        Task<bool> DeleteAccountsHeadManualEntries(long entryid);
        Task<bool> DeleteAccountsHeadManualEntries_SC(long entryid);
        Task<bool> DeleteRemunerationandOtherBenefitsEntries(long entryid);
        Task<bool> DeleteManualEntryForNotes(long entryid);
        Task<bool> DeleteSubHeadsManualEntry(long entryid);
        Task<bool> DeleteVoucher_Adjustment(string vchrno);
        Task<bool> DeleteProfitandLoss(long entryid);
        Task<bool> DeleteBalanceSheetHeads(long entryid);
        Task<bool> DeletCashFlowMainHeads(long entryid);
        Task<bool> DeleteCashFlowReportHeads(long entryid);

        Task GenerateBalanceSheetReportAsync(DateTime dtFrom, DateTime dtTo);
        Task ProcessCapitalWorkInProgress();
        Task SaveProperty();
        Task SaveBalanceSheetToPrint();
        Task PrintIncomeAndExpenseReport();
        Task SaveCashFlowStatementToPrint(DateTime dtFromPrv, DateTime dtToPrv, DateTime dtFromCur, DateTime dtToCur, bool bPrvYear = false);

    }
}

