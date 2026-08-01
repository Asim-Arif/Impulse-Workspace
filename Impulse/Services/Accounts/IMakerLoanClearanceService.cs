using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public interface IMakerLoanClearanceService
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<MakerLoanBalanceModel> GetMakerBalancesAsync(string accNo);
        Task SaveClearanceAsync(MakerLoanClearanceDto dto);
        Task UpdateLTDeductionAmountAsync(string accNo, decimal newDAmount);
    }
}
