using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Data;

namespace CIP.Services
{
    public class AccountReportingAccessService : IAccountReportingAccessService
    {
        private readonly IVouchersDataAccess _dataAccess;
        private readonly IAuditService _auditService;
        private readonly IAccountReportingAccess _accountReportingAccess;
        public AccountReportingAccessService(IVouchersDataAccess dataAccess, IAuditService auditService,IAccountReportingAccess accountReportingAccess)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
            _accountReportingAccess = accountReportingAccess;
        }
        public async Task<List<BOD_Meeting_Payment_List_ViewModel>> GetBODMeetingPaymentList(DateTime DTFrom, DateTime DTTo, int iStatus = 3, int iBankID = 0)
        {
            return await _accountReportingAccess.GetBODMeetingPaymentList(DTFrom, DTTo, iStatus , iBankID );
        }
        public async Task<List<BOD_Meeting_Payment_Detail_ViewModel>> GetBODMeetingPaymentDetail(int EntryID) 
        {
            return await _accountReportingAccess.GetBODMeetingPaymentDetail(EntryID);
        }
        public async Task<List<Online_Payment_Voucher_List_ViewModel>> GetOnlinePaymentVoucherList(DateTime DTFrom, DateTime DTTo, int iBankID = 0, int iStatus = 0, int iVchrType = 0)
        {
            return await _accountReportingAccess.GetOnlinePaymentVoucherList(DTFrom, DTTo, iBankID,iStatus,iVchrType);
        }
        public async Task<List<Trial_Balance_ViewModel>> GetTrialBalanceList(DateTime DTFrom, DateTime DTTo, string Code)
        {
            return await _accountReportingAccess.GetTrialBalanceList(DTFrom, DTTo, Code);
        }
        public async Task<List<Bank_Balance_Statement_ViewModel>> GetBankBalanceStatement() 
        {
            return await _accountReportingAccess.GetBankBalanceStatement();
        }
        public async Task<List<Cheque_Receiving_ViewModel>> GetChequeReceivingData(DateTime DTFrom, DateTime DTTo) 
        {
            return await _accountReportingAccess.GetChequeReceivingData(DTFrom,DTTo);
        }
        public async Task<List<Important_Tasks_ViewModel>> GetImportantTasksData(DateTime DTFrom, DateTime DTTo) 
        {
            return await _accountReportingAccess.GetImportantTasksData(DTFrom, DTTo);
        }
        public async Task<List<Bills_Received_ViewModel>> GetBillsReceivedData(DateTime DTFrom, DateTime DTTo, int SerialFrom, int SerialTo)
        {
            return await _accountReportingAccess.GetBillsReceivedData(DTFrom, DTTo, SerialFrom, SerialTo);
        }
        public async Task<List<NTN_CNIC_ViewModel>> GetNTN_CNICData(string SearchText) 
        {
            return await _accountReportingAccess.GetNTN_CNICData(SearchText);
        }
        public async Task<List<Letter_Dispatch_Register_ViewModel>> GetLetterDispatchData(DateTime DTFrom, DateTime DTTo) 
        {
            return await _accountReportingAccess.GetLetterDispatchData(DTFrom, DTTo);
        }

    }
}
