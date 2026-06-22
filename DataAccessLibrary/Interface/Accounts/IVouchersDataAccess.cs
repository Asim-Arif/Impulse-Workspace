using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Data;
namespace DataAccessLibrary.Interface.Accounts
{    
    public interface IVouchersDataAccess
    {
        Task<List<GenericDropDownModel>> GetBPVPaymentTypes();
        Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "");
        Task<List<ChartOfAccountsModel>> GetTransactionalAccounts(string? strCondition = null);
        Task<List<BankAccountInfo>> GetBankAccounts();
        Task<ChequeInfo> GetNewChequeNumbersByBankAccount(string strAccNo);
        Task<decimal> GetBalance(string strAccNo, DateTime OnDate);
        Task<string> GetNextVchrNo(DateTime OnDate,string VchrType);
        Task<string> GetNextVchrNo_Approval(DateTime OnDate, string VchrType);
        Task<long> GetNextSNo(DateTime OnDate,string strTableName, IDbConnection? existingDb = null, IDbTransaction? existingTrans = null);
        Task<string> ExecuteVoucherSave_Old(VoucherViewModel voucherViewModel);

        Task<string> ExecuteVoucherSave(VoucherViewModel voucherViewModel, IDbConnection? existingDb = null, IDbTransaction? existingTrans = null);
        
        Task<List<VAD_Voucher_DetailViewModel>> GetPendingJVsByAccount(string accNo);
        Task<Voucher_MainSaveResult> ExecuteVoucher_MainSave(Vouchers_MainViewModel voucherViewModel);
        Task<List<SIALDirectorViewModel>> GetDirectorsByBODMeetingDate(DateTime meetingDate);
        Task<int> ExecuteOnlinePaymentVoucherSave(Online_PVModelView header, List<Online_PV_DetailsModelView> details);
        Task<string> ExecuteBODMeetingPaymentVoucherSave(VoucherViewModel voucher, BOD_Meeting_Payment_List_ViewModel BODPLVM, double chqBookNo, double SNo_DebitLine);
        Task<List<GenericDropDownModel>> GetOnlineVoucherPaymentModes();
    }
}
