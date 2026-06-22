using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;
namespace DataAccessLibrary.Interface.Accounts
{    
    public interface IChqBookDetailDataAccess
    {
        Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "");
        Task<List<BankAccountInfo>> GetBankAccounts();
        Task<string> ExecuteChequeBookSave(ChqBookDetailViewModel chqbookViewModel);
        Task<List<ChqBookDetailViewModel>> GetBankChqBooks(string StrAccNo);
        Task<bool> DeleteChqBook(long ChqBookNo);
        Task<List<BankAccountInfo>> GetBankAccountsList();
        Task<bool> SaveOverDraftData(BankAccountInfo bankAccount,string StrAccNo,bool bdefaultbank);
        Task<List<ChqListModel>> GetChqBookChequesList(long chqbookno);
        Task<bool> ChangeChqStatus(ChqListModel chqs, string chqno, double sno, int iType, long INextSno);
    }
}
