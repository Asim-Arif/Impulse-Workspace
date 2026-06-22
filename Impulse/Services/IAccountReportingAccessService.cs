using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;

namespace CIP.Services
{
    public interface IAccountReportingAccessService
    {
        Task<List<BOD_Meeting_Payment_List_ViewModel>> GetBODMeetingPaymentList(DateTime DTFrom, DateTime DTTo, int iStatus = 3, int iBankID = 0);
        Task<List<BOD_Meeting_Payment_Detail_ViewModel>> GetBODMeetingPaymentDetail(int EntryID);
        Task<List<Online_Payment_Voucher_List_ViewModel>> GetOnlinePaymentVoucherList(DateTime DTFrom, DateTime DTTo, int iBankID = 0, int iStatus = 0, int iVchrType = 0);
        Task<List<Trial_Balance_ViewModel>> GetTrialBalanceList(DateTime DTFrom, DateTime DTTo, string Code);
        Task<List<Bank_Balance_Statement_ViewModel>> GetBankBalanceStatement();
        Task<List<Cheque_Receiving_ViewModel>> GetChequeReceivingData(DateTime DTFrom, DateTime DTTo);
        Task<List<Important_Tasks_ViewModel>> GetImportantTasksData(DateTime DTFrom, DateTime DTTo);
        Task<List<Bills_Received_ViewModel>> GetBillsReceivedData(DateTime DTFrom, DateTime DTTo, int SerialFrom, int SerialTo);
        Task<List<NTN_CNIC_ViewModel>> GetNTN_CNICData(string SearchText);
        Task<List<Letter_Dispatch_Register_ViewModel>> GetLetterDispatchData(DateTime DTFrom, DateTime DTTo);
    }
}
