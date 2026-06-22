using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;

namespace Impulse.Services
{
    public interface IVoucherService
    {
        Task<string> SaveVoucherAsync(VoucherViewModel voucherViewModel);
        Task<Voucher_MainSaveResult> SaveVoucherMainAsync(Vouchers_MainViewModel voucherViewModel);
        Task<List<GenericDropDownModel>> GetBPVPaymentTypes();
        Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "");
        Task<List<ChartOfAccountsModel>> GetTransactionalAccounts(string? strCondition = null);
        Task<List<BankAccountInfo>> GetBankAccounts();
        Task<ChequeInfo> GetNewChequeNumbersByBankAccount(string strAccNo);
        Task<decimal> GetBalance(string strAccNo, DateTime OnDate);
        Task<string> GetNextVchrNo(DateTime OnDate, string VchrType);
        Task<string> GetNextVchrNo_Approval(DateTime OnDate, string VchrType);
        Task<long> GetNextSNo(DateTime OnDate, string strTableName);
        Task<List<VAD_Voucher_DetailViewModel>> GetPendingJVsByAccount(string accNo);
        Task<List<SIALDirectorViewModel>> GetDirectorsByBODMeetingDate(DateTime meetingDate);
        Task<int> SaveOnlinePaymentVoucher(Online_PVModelView header, List<Online_PV_DetailsModelView> details);
        Task<string> SaveBODMeetingPaymentVoucher(VoucherViewModel voucher, BOD_Meeting_Payment_List_ViewModel BODPLVM, double chqBookNo, double SNo_DebitLine);
        Task<List<GenericDropDownModel>> GetOnlineVoucherPaymentModes();
    }
}
