using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public interface IMakerLoanTransferService
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<MakerLoanBalanceModel> GetMakerBalancesAsync(string accNo);
        Task SaveTransferAsync(MakerLoanTransferDto dto);
    }
}
