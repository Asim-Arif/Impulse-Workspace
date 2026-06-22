 using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;
using static DataAccessLibrary.Models.ViewModels.Accounts.MainHeadsModel;

public interface IFinancialStatementDataAccess
{
    Task<List<MainHeadsModel>> GetMainHeadsList();
    Task<List<SubHeadsModel>> GetSubHeadList();
    Task<List<SubHeadsModel>> GetMainHeads();
    Task<List<GroupsAccountsModel>> GetGroupsAccountsList();
    Task<List<SubHeadsListModel>> GetSubHeads(long MainGroup_refID);
    Task<string> GetNextVchrNo_Adjustment(DateTime OnDate, string VchrType);
    Task<List<MainHeadsListModel>> GetAllHeads();



    Task<bool> SaveMainHead(MainHeadsModel mainhead);
    Task<bool> EditMainHead(MainHeadsModel mainhead, long entryid);
    Task<bool> DeleteMainHead(long entryid);

    Task<bool> SaveSubHead(SubHeadsModel subhead);
    Task<bool> EditSubHead(SubHeadsModel subhead, long entryid, int itype);
    Task<bool> DeleteSubHead(long entryid);

    Task<List<SubHeadAccountDetailModel>> GetAccountGroup_HeadsDetail(long AccGroup_RefID);
    Task<bool> SaveSubHeadDetail(SubHeadAccountDetailModel subhead);
    Task<bool> EditSubHeadDetailEntry(SubHeadAccountDetailModel subhead, long entryid,int _itype);
    Task<bool> DeleteSubHeadEntry(long entryid);


    Task<bool> SaveGroupsAccounts(GroupsAccountsModel groupsaccounts);
    Task<bool> EditGroupsAccounts(GroupsAccountsModel groupsaccounts, long entryid);
    Task<bool> DeleteGroupsAccounts(long entryid);


    Task<List<AccountsHeadManualEntryModel>> GetAccountsHeadManualEntries();
    Task<bool> SaveAccountsHeadManualEntries(AccountsHeadManualEntryModel groupsaccounts);
    Task<bool> EditAccountsHeadManualEntries(AccountsHeadManualEntryModel mainhead, long entryid);
    Task<bool> DeleteAccountsHeadManualEntries(long entryid);


    Task<List<AccountsHeadManualEntry_SC_Model>> GetAccountsHeadManualEntries_SC();
    Task<bool> SaveAccountsHeadManualEntries_SC(AccountsHeadManualEntry_SC_Model groupsaccounts);
    Task<bool> EditAccountsHeadManualEntries_SC(AccountsHeadManualEntry_SC_Model mainhead, long entryid);
    Task<bool> DeleteAccountsHeadManualEntries_SC(long entryid);

    Task<List<RemunerationandOtherBenefitsModel>> GetRemunerationandOtherBenefitsEntries();
    Task<bool> SaveRemunerationandOtherBenefitsEntries(RemunerationandOtherBenefitsModel groupsaccounts);
    Task<bool> EditRemunerationandOtherBenefitsEntries(RemunerationandOtherBenefitsModel mainhead, long entryid);
    Task<bool> DeleteRemunerationandOtherBenefitsEntries(long entryid);


    Task<List<ManualEntryForNotesModel>> GetManualEntryForNotesList();
    Task<bool> SaveManualEntryForNotes(ManualEntryForNotesModel groupsaccounts);
    Task<bool> EditManualEntryForNotes(ManualEntryForNotesModel mainhead, long entryid);
    Task<bool> DeleteManualEntryForNotes(long entryid);


    Task<List<SubHeadsManualEntryModel>> GetSubHeadsManualEntriesList();
    Task<bool> SaveSubHeadsManualEntries(SubHeadsManualEntryModel groupsaccounts);
    Task<bool> EditSubHeadsManualEntries(SubHeadsManualEntryModel mainhead, long entryid);
    Task<bool> DeleteSubHeadsManualEntry(long entryid);


    Task<List<Vouchers_Adjustment_Transaction_Model>> GetTransactionData_Adjustment(DateTime DTFrom, DateTime DTTo);
    Task<bool> DeleteVoucher_Adjustment(string vchrno);
    Task<string> SaveAdjustmentVoucher(Vouchers_Main_Model voucheradjustmentmodel);


    Task<List<ProfitandLossHeadsDetailToLoadModel>> GetProfitandLossHeadsList();
    Task<string> SaveProfitandLossHeads(ProfitandLossHeads_Main_Model profitandlossmain);
    Task<bool> EditProfitandLossHeads(ProfitandLossHeads_Main_Model mainhead, long entryid);
    Task<bool> DeleteProfitandLoss(long entryid);


    Task<List<BalanceSheetHeadsDetailToLoadModel>> GetBalanceSheetHeadsList();
    Task<string> SaveBalanceSheetHeads(BalanceSheet_Main_Model balancesheet);
    Task<bool> EditBalanceSheetHeads(BalanceSheet_Main_Model balancesheet, long entryid);
    Task<bool> DeleteBalanceSheetHeads(long entryid);

    Task<List<FinancialStatementHeadsModel>> GetFinancialStatementHeadsList();
    Task<string> SaveFinancialStatementHeads(string username, string machinename);
    Task<bool> EditFinancialStatementHeads(FinancialStatementHeadsModel mainhead, int itype, long entryid);

    Task<List<CashFlowMainHeadsModel>> GetCashFlowMainHeadsList();
    Task<List<CashFlowType_DD_Model>> GetCashFlowType_DD();
    Task<List<CashFlowMain_DD_Model>> GetCashFlowMainHeads_DD();
    Task<bool> SaveCashFlowMainHeads(CashFlowMainHeadsModel groupsaccounts);
    Task<bool> EditCashFlowMainHeads(CashFlowMainHeadsModel mainhead, int itype, long entryid);
    Task<bool> DeletCashFlowMainHeads(long entryid);

    Task<List<CashFlowReportHeadsDetailToLoadModel>> GetCashFlowReportHeadsList();
    Task<List<CashFlowHeadsListModel>> GetCashFlowHeadsList_DD();
    Task<string> SaveCashFlowReportHeads(CashFlowReportHeads_Main_Model cashflowmain);
    Task<bool> DeleteCashFlowReportHeads(long entryid);
    Task<bool> EditCashFlowReportHeads(CashFlowReportHeads_Main_Model mainhead, long entryid);


    Task<List<PrintAttachmentsModel>> GetNotesAttachmentList();
    Task<int> SaveNotesAttachments(PrintAttachmentsModel groupsaccounts);


    Task GenerateBalanceSheetReportAsync(DateTime dtFrom, DateTime dtTo);
    Task GenerateBalanceSheetReportAsync_Prv(DateTime dtFrom, DateTime dtTo);
    Task ProcessCapitalWorkInProgress();
    Task SaveProperty();
    Task SaveBalanceSheetToPrint();
    Task PrintIncomeAndExpenseReport();

    Task SaveCashFlowStatementToPrint(DateTime dtFromPrv, DateTime dtToPrv, DateTime dtFromCur, DateTime dtToCur, bool bPrvYear = false);

}
