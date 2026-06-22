using CIP.Pages.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Server.Kestrel.Core.Features;
using System.Data;

namespace CIP.Services
{
    public class FinancialStatementService : IFinancialStatementService
    {
        private readonly IFinancialStatementDataAccess _dataAccess;
        private readonly IAuditService _auditService;
        private readonly IVoucherService _voucherService;
        private readonly IDBHelperService _dbhelperService;
        private readonly IAccountReportingAccess _accountreporting;

        public FinancialStatementService(IFinancialStatementDataAccess dataAccess, IAuditService auditService,
            IVoucherService voucherService, IDBHelperService dbhelperService, IAccountReportingAccess accountreporting)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
            _voucherService = voucherService;
            _dbhelperService = dbhelperService;
            _accountreporting = accountreporting;
        }
        public async Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "")
        {
            return await _voucherService.GetValuesForDropDown(strTableName, strFieldName_ID, strFieldName_Description, strCondition);
        }
        public async Task<List<MainHeadsModel>> GetMainHeadsList()
        {
            return await _dataAccess.GetMainHeadsList();
        }
        public Task<List<SubHeadsModel>> GetSubHeadList()
        {
            return _dataAccess.GetSubHeadList();
        }
        public Task<List<GroupsAccountsModel>> GetGroupsAccountsList()
        {
            return _dataAccess.GetGroupsAccountsList();
        }

        public async Task<List<SubHeadsModel>> GetMainHeads()
        {
            return await _dataAccess.GetMainHeads();
        }
        public Task<List<SubHeadAccountDetailModel>> GetAccountGroup_HeadsDetail(long AccGroup_RefID)
        {
            return _dataAccess.GetAccountGroup_HeadsDetail(AccGroup_RefID);
        }
        public Task<List<AccountsHeadManualEntryModel>> GetAccountsHeadManualEntries()
        {
            return _dataAccess.GetAccountsHeadManualEntries();
        }
        public async Task<List<SubHeadsListModel>> GetSubHeads(long MainGroup_refID)
        {
            return await _dataAccess.GetSubHeads(MainGroup_refID);
        }
        public Task<List<AccountsHeadManualEntry_SC_Model>> GetAccountsHeadManualEntries_SC()
        {
            return _dataAccess.GetAccountsHeadManualEntries_SC();
        }
        public Task<List<RemunerationandOtherBenefitsModel>> GetRemunerationandOtherBenefitsEntries()
        {
            return _dataAccess.GetRemunerationandOtherBenefitsEntries();
        }
        public Task<List<ManualEntryForNotesModel>> GetManualEntryForNotesList()
        {
            return _dataAccess.GetManualEntryForNotesList();
        }
        public Task<List<SubHeadsManualEntryModel>> GetSubHeadsManualEntriesList()
        {
            return _dataAccess.GetSubHeadsManualEntriesList();
        }
        public Task<List<Vouchers_Adjustment_Transaction_Model>> GetTransactionData_Adjustment(DateTime DTFrom, DateTime DTTo)
        {
            return _dataAccess.GetTransactionData_Adjustment(DTFrom, DTTo);
        }
        public async Task<string> GetNextVchrNo_Adjustment(DateTime OnDate, string VoucherType)
        {
            string nextVoucherNo = await _dataAccess.GetNextVchrNo_Adjustment(OnDate, VoucherType);
            return nextVoucherNo;
        }
        public async Task<List<MainHeadsListModel>> GetAllHeads()
        {
            return await _dataAccess.GetAllHeads();
        }
        public Task<List<ProfitandLossHeadsDetailToLoadModel>> GetProfitandLossHeadsList()
        {
            return _dataAccess.GetProfitandLossHeadsList();
        }
        public Task<List<BalanceSheetHeadsDetailToLoadModel>> GetBalanceSheetHeadsList()
        {
            return _dataAccess.GetBalanceSheetHeadsList();
        }
        public Task<List<FinancialStatementHeadsModel>> GetFinancialStatementHeadsList()
        {
            return _dataAccess.GetFinancialStatementHeadsList();
        }
        public Task<List<PrintAttachmentsModel>> GetNotesAttachmentList()
        {
            return _dataAccess.GetNotesAttachmentList();
        }
        public Task<List<CashFlowMainHeadsModel>> GetCashFlowMainHeadsList()
        {
            return _dataAccess.GetCashFlowMainHeadsList();
        }
        public Task<List<CashFlowType_DD_Model>> GetCashFlowType_DD()
        {
            return _dataAccess.GetCashFlowType_DD();
        }
        public Task<List<CashFlowMain_DD_Model>> GetCashFlowMainHeads_DD()
        {
            return _dataAccess.GetCashFlowMainHeads_DD();
        }
        public Task<List<CashFlowReportHeadsDetailToLoadModel>> GetCashFlowReportHeadsList()
        {
            return _dataAccess.GetCashFlowReportHeadsList();
        }
        public async Task<List<CashFlowHeadsListModel>> GetCashFlowHeadsList_DD()
        {
            return await _dataAccess.GetCashFlowHeadsList_DD();
        }


        ////////////// SAVE
        public async Task<bool> SaveMainHead(MainHeadsModel mainhead)
        {
            return await _dataAccess.SaveMainHead(mainhead);
        }
        public async Task<bool> SaveSubHead(SubHeadsModel subhead)
        {
            return await _dataAccess.SaveSubHead(subhead);
        }
        public async Task<bool> SaveSubHeadDetail(SubHeadAccountDetailModel subhead)
        {
            return await _dataAccess.SaveSubHeadDetail(subhead);
        }
        public async Task<bool> SaveGroupsAccounts(GroupsAccountsModel groupaccount)
        {
            if (groupaccount != null)
            {
                long sortno = _dbhelperService.getSingleIntValue("MAX(SortNo)", "GroupsForExpenseReport");
                sortno = sortno + 1;
                groupaccount.SortNo = sortno.ToString();
                groupaccount.UserName = _auditService.GetCurrentUserName();
                groupaccount.MachineName = _auditService.GetClientIpAddress();
            }

            return await _dataAccess.SaveGroupsAccounts(groupaccount);
        }
        public async Task<bool> SaveAccountsHeadManualEntries(AccountsHeadManualEntryModel groupaccount)
        {
            if (groupaccount != null)
            {
                groupaccount.UserName =  _auditService.GetCurrentUserName();
                groupaccount.MachineName = _auditService.GetClientIpAddress();
            }
            return await _dataAccess.SaveAccountsHeadManualEntries(groupaccount);
        }
        public async Task<bool> SaveAccountsHeadManualEntries_SC(AccountsHeadManualEntry_SC_Model groupaccount)
        {
            if (groupaccount != null)
            {
                groupaccount.UserName = _auditService.GetCurrentUserName();
                groupaccount.MachineName = _auditService.GetClientIpAddress();
            }
            return await _dataAccess.SaveAccountsHeadManualEntries_SC(groupaccount);
        }
        public async Task<bool> SaveRemunerationandOtherBenefitsEntries(RemunerationandOtherBenefitsModel groupaccount)
        {
            if (groupaccount != null)
            {
                groupaccount.UserName =  _auditService.GetCurrentUserName();
                groupaccount.MachineName = _auditService.GetClientIpAddress();
            }
            return await _dataAccess.SaveRemunerationandOtherBenefitsEntries(groupaccount);
        }
        public async Task<bool> SaveManualEntryForNotes(ManualEntryForNotesModel groupaccount)
        {
            if (groupaccount != null)
            {
                groupaccount.UserName = _auditService.GetCurrentUserName();
                groupaccount.MachineName = _auditService.GetClientIpAddress();
            }
            return await _dataAccess.SaveManualEntryForNotes(groupaccount);
        }
        public async Task<bool> SaveSubHeadsManualEntries(SubHeadsManualEntryModel groupaccount)
        {
            if (groupaccount != null)
            {
                long sortno = _dbhelperService.getSingleIntValue("MAX(SrNo)", "SubHeadsManualEntry_Accounts");
                sortno = sortno + 1;
                groupaccount.SrNo = sortno;

                groupaccount.UserName = _auditService.GetCurrentUserName();
                groupaccount.MachineName = _auditService.GetClientIpAddress();
            }
            return await _dataAccess.SaveSubHeadsManualEntries(groupaccount);
        }
        public async Task<string> SaveAdjustmentVoucher(Vouchers_Main_Model voucher)
        {
            if (voucher != null)
            {
                voucher.UserName =  _auditService.GetCurrentUserName();
                voucher.MachineName = _auditService.GetClientIpAddress();
            }

            if (voucher.LineItems != null && voucher.LineItems.Any())
            {
                foreach (var detail in voucher.LineItems)
                {
                    decimal Balance = await _accountreporting.GetAccountOpeningBalance(detail.AccNo, voucher.DT);
                    detail.Balance = Balance - detail.Debit - detail.Credit;
                }
            }
            return await _dataAccess.SaveAdjustmentVoucher(voucher);
        }
        public async Task<string> SaveProfitandLossHeads(ProfitandLossHeads_Main_Model main)
        {
            return await _dataAccess.SaveProfitandLossHeads(main);
        }
        public async Task<string> SaveBalanceSheetHeads(BalanceSheet_Main_Model main)
        {
            return await _dataAccess.SaveBalanceSheetHeads(main);
        }
        public async Task<string> SaveFinancialStatementHeads(string username, string machinename)
        {
            return await _dataAccess.SaveFinancialStatementHeads(username, machinename);
        }
        public async Task<int> SaveNotesAttachments(PrintAttachmentsModel groupaccount)
        {
            return await _dataAccess.SaveNotesAttachments(groupaccount);
        }
        public async Task<bool> SaveCashFlowMainHeads(CashFlowMainHeadsModel groupaccount)
        {
            return await _dataAccess.SaveCashFlowMainHeads(groupaccount);
        }
        public async Task<string> SaveCashFlowReportHeads(CashFlowReportHeads_Main_Model main)
        {
            return await _dataAccess.SaveCashFlowReportHeads(main);
        }


        ////////////// EDIT
        public async Task<bool> EditMainHead(MainHeadsModel mainhead, long entryid)
        {
            return await _dataAccess.EditMainHead(mainhead, entryid);
        }
        public async Task<bool> EditSubHead(SubHeadsModel subhead, long entryid, int itype)
        {
            return await _dataAccess.EditSubHead(subhead, entryid, itype);
        }
        public async Task<bool> EditSubHeadDetailEntry(SubHeadAccountDetailModel subhead, long entryid, int _itype)
        {
            return await _dataAccess.EditSubHeadDetailEntry(subhead, entryid, _itype);
        }
        public async Task<bool> EditGroupsAccounts(GroupsAccountsModel subhead, long entryid)
        {
            return await _dataAccess.EditGroupsAccounts(subhead, entryid);
        }
        public async Task<bool> EditAccountsHeadManualEntries(AccountsHeadManualEntryModel subhead, long entryid)
        {
            return await _dataAccess.EditAccountsHeadManualEntries(subhead, entryid);
        }
        public async Task<bool> EditAccountsHeadManualEntries_SC(AccountsHeadManualEntry_SC_Model subhead, long entryid)
        {
            return await _dataAccess.EditAccountsHeadManualEntries_SC(subhead, entryid);
        }
        public async Task<bool> EditRemunerationandOtherBenefitsEntries(RemunerationandOtherBenefitsModel subhead, long entryid)
        {
            return await _dataAccess.EditRemunerationandOtherBenefitsEntries(subhead, entryid);
        }
        public async Task<bool> EditManualEntryForNotes(ManualEntryForNotesModel subhead, long entryid)
        {
            return await _dataAccess.EditManualEntryForNotes(subhead, entryid);
        }
        public async Task<bool> EditSubHeadsManualEntries(SubHeadsManualEntryModel subhead, long entryid)
        {
            return await _dataAccess.EditSubHeadsManualEntries(subhead, entryid);
        }
        public async Task<bool> EditProfitandLossHeads(ProfitandLossHeads_Main_Model main, long entryid)
        {
            return await _dataAccess.EditProfitandLossHeads(main, entryid);
        }
        public async Task<bool> EditBalanceSheetHeads(BalanceSheet_Main_Model main, long entryid)
        {
            return await _dataAccess.EditBalanceSheetHeads(main, entryid);
        }
        public async Task<bool> EditFinancialStatementHeads(FinancialStatementHeadsModel main, int itype, long entryid)
        {
            return await _dataAccess.EditFinancialStatementHeads(main, itype, entryid);
        }
        public async Task<bool> EditCashFlowMainHeads(CashFlowMainHeadsModel main, int itype, long entryid)
        {
            return await _dataAccess.EditCashFlowMainHeads(main, itype, entryid);
        }
        public async Task<bool> EditCashFlowReportHeads(CashFlowReportHeads_Main_Model main, long entryid)
        {
            return await _dataAccess.EditCashFlowReportHeads(main, entryid);
        }


        ////////////// DELETE
        public async Task<bool> DeleteMainHead(long entryid)
        {
            long icount = await _dbhelperService.getSingleIntValueasync("COUNT(*)", "GroupsForFinancialReport_Sub", "WHERE Type=" + entryid + "");
            if (icount > 0)
            {
                throw new Exception("Cannot delete Main Head as it has related Sub Heads.");
            }
            return await _dataAccess.DeleteMainHead(entryid);
        }
        public async Task<bool> DeleteSubHead(long entryid)
        {
            long icount = await _dbhelperService.getSingleIntValueasync("COUNT(*)", "AccountGroupsForBalanceSheet", "WHERE AccGroup_RefID=" + entryid + "");
            if (icount > 0)
            {
                throw new Exception("Cannot delete Head as it has related accounts head.");
            }
            return await _dataAccess.DeleteSubHead(entryid);
        }
        public async Task<bool> DeleteSubHeadEntry(long entryid)
        {
            return await _dataAccess.DeleteSubHeadEntry(entryid);
        }
        public async Task<bool> DeleteGroupsAccounts(long entryid)
        {
            return await _dataAccess.DeleteGroupsAccounts(entryid);
        }
        public async Task<bool> DeleteAccountsHeadManualEntries(long entryid)
        {
            return await _dataAccess.DeleteAccountsHeadManualEntries(entryid);
        }
        public async Task<bool> DeleteAccountsHeadManualEntries_SC(long entryid)
        {
            return await _dataAccess.DeleteAccountsHeadManualEntries_SC(entryid);
        }
        public async Task<bool> DeleteRemunerationandOtherBenefitsEntries(long entryid)
        {
            return await _dataAccess.DeleteRemunerationandOtherBenefitsEntries(entryid);
        }
        public async Task<bool> DeleteManualEntryForNotes(long entryid)
        {
            return await _dataAccess.DeleteManualEntryForNotes(entryid);
        }
        public async Task<bool> DeleteSubHeadsManualEntry(long entryid)
        {
            return await _dataAccess.DeleteSubHeadsManualEntry(entryid);
        }
        public async Task<bool> DeleteVoucher_Adjustment(string vchrno)
        {
            return await _dataAccess.DeleteVoucher_Adjustment(vchrno);
        }
        public async Task<bool> DeleteProfitandLoss(long entryid)
        {
            return await _dataAccess.DeleteProfitandLoss(entryid);
        }
        public async Task<bool> DeleteBalanceSheetHeads(long entryid)
        {
            return await _dataAccess.DeleteBalanceSheetHeads(entryid);
        }
        public async Task<bool> DeletCashFlowMainHeads(long entryid)
        {
            return await _dataAccess.DeletCashFlowMainHeads(entryid);
        }
        public async Task<bool> DeleteCashFlowReportHeads(long entryid)
        {
            return await _dataAccess.DeleteCashFlowReportHeads(entryid);
        }

        public async Task GenerateBalanceSheetReportAsync(DateTime dtFrom, DateTime dtTo)
        {
            await _dataAccess.GenerateBalanceSheetReportAsync(dtFrom, dtTo);
            await _dataAccess.GenerateBalanceSheetReportAsync_Prv(dtFrom, dtTo);
        }
        public async Task ProcessCapitalWorkInProgress()
        {
            await _dataAccess.ProcessCapitalWorkInProgress();
        }
        public async Task SaveProperty()
        {
            await _dataAccess.SaveProperty();   
        }
        public async Task SaveBalanceSheetToPrint()
        {
            await _dataAccess.SaveBalanceSheetToPrint();
        }
        public async Task PrintIncomeAndExpenseReport()
        {
            await _dataAccess.PrintIncomeAndExpenseReport();
        }

        public async Task SaveCashFlowStatementToPrint(DateTime dtFromPrv, DateTime dtToPrv, DateTime dtFromCur, DateTime dtToCur, bool bPrvYear = false)
        {
            await _dataAccess.SaveCashFlowStatementToPrint(dtFromPrv, dtToPrv, dtFromCur, dtToCur, bPrvYear);
        }

    }
}