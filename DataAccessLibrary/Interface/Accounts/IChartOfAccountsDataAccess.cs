using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;

namespace DataAccessLibrary.Interface.Accounts
{
    public interface IChartOfAccountsDataAccess
    {
        Task<List<ChartOfAccountsModel>> GetChartOfAccountsList(bool showinactive);
        Task<List<ChartOfAccountsModel>> GetAccountsHeads();
        Task<List<ChartOfAccountsModel>> GetSubAccOfAccounts(string acctype);
        Task<string> GetNextAccountNumberAsync(string acctype, string accountname, Boolean isparent, string subaccof);
        Task SaveNewAccount(ChartOfAccountsModel newaccount);
        Task EditAccount(ChartOfAccountsModel newaccount,string AccNo);
        
        Task DeleteAccountAsync(string AccNo);
        Task UpdateAccountStatusAsync(string AccNo,Boolean bActive);

        Task<List<CashFlowHeadsModel>> GetCashFlowHeads();
        Task UpdateCashFlowHeads(ChartOfAccountsModel newaccount, string AccNo, long CFHFS_RefID = 0, int itype = 0);
        Task DeleteAccountGroupsForBalancesheet(string AccNo);
        Task SaveAccountGroupsForBalancesheet(ChartOfAccountsModel coa, string accno, long AccGroup_RefID, int itype);

        Task<List<NotesToAccountsModel>> GetNotesToAccountsHeads();
    }
}