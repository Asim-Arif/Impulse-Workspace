using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Accounts
{
    public interface IMakerLoanClearanceDataAccess
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<MakerLoanBalanceModel> GetMakerBalancesAsync(string accNo);
        Task SaveClearanceAsync(MakerLoanClearanceDto dto, string userName, string machineName);
        Task UpdateLTDeductionAmountAsync(string accNo, decimal newDAmount);
    }
}
