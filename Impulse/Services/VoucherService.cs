using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Data;

namespace CIP.Services
{
    public class VoucherService : IVoucherService
    {
        private readonly IVouchersDataAccess _dataAccess;
        private readonly IAuditService _auditService;
        public VoucherService(IVouchersDataAccess dataAccess, IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
        }
        public async Task<string> GetNextVchrNo(DateTime OnDate, string VoucherType)
        { 
            string nextVoucherNo = await _dataAccess.GetNextVchrNo(OnDate, VoucherType);
            return nextVoucherNo;
        }
        public async Task<string> GetNextVchrNo_Approval(DateTime OnDate, string VoucherType)
        {
            string nextVoucherNo = await _dataAccess.GetNextVchrNo_Approval(OnDate, VoucherType);
            return nextVoucherNo;
        }
        public async Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "")
        {            
            return await _dataAccess.GetValuesForDropDown(strTableName, strFieldName_ID, strFieldName_Description, strCondition);
        }
        public async Task<List<GenericDropDownModel>> GetBPVPaymentTypes()
        {
            return await _dataAccess.GetBPVPaymentTypes();
        }
        public async Task<List<GenericDropDownModel>> GetOnlineVoucherPaymentModes()
        {
            return await _dataAccess.GetOnlineVoucherPaymentModes();
        }
        public async Task<List<ChartOfAccountsModel>> GetTransactionalAccounts(string? strCondition = null)
        {
            return await _dataAccess.GetTransactionalAccounts(strCondition);
        }
        public async Task<List<BankAccountInfo>> GetBankAccounts()
        { 
            return await _dataAccess.GetBankAccounts();
        }
        public async Task<ChequeInfo> GetNewChequeNumbersByBankAccount(string strAccNo)
        { 
            return await _dataAccess.GetNewChequeNumbersByBankAccount(strAccNo);
        }
        public async Task<decimal> GetBalance(string strAccNo, DateTime OnDate)
        { 
            return await _dataAccess.GetBalance(strAccNo, OnDate); 
        }
        public async Task<long> GetNextSNo(DateTime OnDate, string strTableName)
        { 
            return await _dataAccess.GetNextSNo(OnDate, strTableName);
        }
        public async Task<List<VAD_Voucher_DetailViewModel>> GetPendingJVsByAccount(string accNo) 
        {
            return await _dataAccess.GetPendingJVsByAccount(accNo);
        }
        public async Task<string> SaveVoucherAsync(VoucherViewModel voucher)
        {
            voucher.UserName = _auditService.GetCurrentUserName();
            voucher.MachineName=_auditService.GetClientIpAddress();

            long SNo = await _dataAccess.GetNextSNo(voucher.DT, "Vouchers");
            long ChequeSNo = 0;//await _dataAccess.GetNextSNo(voucher.DT, "Cheque");
            if (voucher.ChequeDetails != null)
            {
                ChequeSNo = await _dataAccess.GetNextSNo(voucher.DT, "Cheque");
                voucher.ChequeDetails.SNo = ChequeSNo;                
            }
            decimal dLargestDebit = 0;
            decimal Balance = 0;
            foreach (var line in voucher.LineItems)
            {
                Balance = 0;//Balance will be updated in [dbo].[InsertVoucherLineItems_SP] //await GetBalance(line.AccNo, voucher.DT);
                //Balance = Balance + line.Debit - line.Credit;
                line.SNo = SNo;
                SNo += 1;
                line.Balance = Balance; //passing 0 as balance will be updated in SP
                line.CSNo = 0;
                line.VDate = voucher.DT;
                if (voucher.ChequeDetails != null)  //this is to make sure that if it's Bank Payment voucher then update CSNo
                {
                    if (line.AccNo == voucher.ChequeDetails.BankAccNo)
                        line.CSNo = (long)ChequeSNo;

                    if (line.Debit > dLargestDebit) //this is to get the Cheque Account No.
                    {
                        dLargestDebit = line.Debit;
                        voucher.ChequeDetails.AccNo = line.AccNo;
                    }
                }
            }
            string strVchrNo=await _dataAccess.ExecuteVoucherSave(voucher);
            return strVchrNo;
        }
        public async Task<Voucher_MainSaveResult> SaveVoucherMainAsync(Vouchers_MainViewModel voucher)
        {
            voucher.UserName = _auditService.GetCurrentUserName();
            voucher.MachineName = _auditService.GetClientIpAddress();

            decimal Balance = 0;
            foreach (var line in voucher.LineItems)
            {
                Balance = await GetBalance(line.AccNo, voucher.DT);                
                line.Balance = Balance; //passing 0 as balance will be updated in SP
                line.CSNo = 0;                                
            }
            Voucher_MainSaveResult voucher_MainSaveResult = await _dataAccess.ExecuteVoucher_MainSave(voucher);
            return voucher_MainSaveResult;
        }
        public async Task<List<SIALDirectorViewModel>> GetDirectorsByBODMeetingDate(DateTime meetingDate)
        {
            return await _dataAccess.GetDirectorsByBODMeetingDate(meetingDate);
        }
        
        public async Task<int> SaveOnlinePaymentVoucher(Online_PVModelView header, List<Online_PV_DetailsModelView> details) 
        {
            return await _dataAccess.ExecuteOnlinePaymentVoucherSave(header, details);
        }
        public async Task<string> SaveBODMeetingPaymentVoucher(VoucherViewModel voucher, BOD_Meeting_Payment_List_ViewModel BODPLVM, double chqBookNo,double SNo_DebitLine ) 
        {
            return await _dataAccess.ExecuteBODMeetingPaymentVoucherSave(voucher, BODPLVM, chqBookNo, SNo_DebitLine);
        }
    }
}
