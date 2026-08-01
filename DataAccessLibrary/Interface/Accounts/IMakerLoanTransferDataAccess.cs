using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Accounts
{
    public interface IMakerLoanTransferDataAccess
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<MakerLoanBalanceModel> GetMakerBalancesAsync(string accNo);
        Task SaveTransferAsync(MakerLoanTransferDto dto, string userName, string machineName);
    }
}
