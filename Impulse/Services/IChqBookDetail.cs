using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;

namespace CIP.Services
{
    public interface IChqBookDetail
    {
        Task<string> SaveChqBook(ChqBookDetailViewModel ChqBookModel);
        Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "");
        Task<List<BankAccountInfo>> GetBankAccounts();
        Task<List<ChqBookDetailViewModel>> GetBankChqBooks(string StrAccNo);
        Task<List<BankAccountInfo>> GetBankAccountsList();
        Task<bool> DeleteChqBook(long ChqBookNo);
        Task<bool> SaveOverDraftData(BankAccountInfo bankaccount, string StrBankAccNo, bool bdefaultbank);
        Task<List<ChqListModel>> GetChqBookChequesList(long chqbookno);
        Task<bool> ChangeChqStatus(ChqListModel chqs, string chqno, double sno, int iType, long INextSno);
    }
}
