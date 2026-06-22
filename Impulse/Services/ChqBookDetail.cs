using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Interface.Accounts;

namespace Impulse.Services
{
    public class ChqBookDetail : IChqBookDetail
    {
        private readonly IChqBookDetailDataAccess _dataAccess;
        private readonly IAuditService _auditService;
        private readonly IVoucherService _voucherService;

        public ChqBookDetail(IChqBookDetailDataAccess dataAccess, IAuditService auditService, IVoucherService voucherService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
            _voucherService = voucherService;
        }
        public async Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "")
        {
            return await _dataAccess.GetValuesForDropDown(strTableName, strFieldName_ID, strFieldName_Description, strCondition);
        }
        public async Task<List<BankAccountInfo>> GetBankAccounts()
        {
            return await _dataAccess.GetBankAccounts();
        }
        public async Task<string> SaveChqBook(ChqBookDetailViewModel chqbook)
        {
            chqbook.UserName = _auditService.GetCurrentUserName();
            chqbook.MachineName = _auditService.GetClientIpAddress();

            string strVchrNo = await _dataAccess.ExecuteChequeBookSave(chqbook);
            return strVchrNo;
        }

        public async Task<List<ChqBookDetailViewModel>> GetBankChqBooks(string StrAccNo = "")
        {
            return await _dataAccess.GetBankChqBooks(StrAccNo);
        }

        public async Task<bool> DeleteChqBook(long chqbookno)
        {
            return await _dataAccess.DeleteChqBook(chqbookno);
        }

        public async Task<List<BankAccountInfo>> GetBankAccountsList()
        {
            return await _dataAccess.GetBankAccountsList();
        }

        public async Task<bool> SaveOverDraftData(BankAccountInfo bankaccount, string StrAccNo, bool bdefaultbank)
        {
            return await _dataAccess.SaveOverDraftData(bankaccount, StrAccNo, bdefaultbank);

        }
        public async Task<List<ChqListModel>> GetChqBookChequesList(long chqbookno)
        {
            return await _dataAccess.GetChqBookChequesList(chqbookno);
        }
        public async Task<bool> ChangeChqStatus(ChqListModel chqs, string chqno, double sno, int iType, long INextSno)
        {
            long ChequeSNo = 0;//await _dataAccess.GetNextsSNo(voucher.DT, "Cheque");
            DateTime dateOnly = DateTime.Now.Date ;
            ChequeSNo = await _voucherService.GetNextSNo(dateOnly, "Cheque");
            INextSno = ChequeSNo;
            return await _dataAccess.ChangeChqStatus(chqs, chqno, sno, iType, INextSno);       
        }

    }
}