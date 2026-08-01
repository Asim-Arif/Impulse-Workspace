using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public interface IMakerLongTermLoanService
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<decimal> GetMakerBalanceAsync(string accNo);
        Task<List<MakerMasterPoModel>> GetMakerMasterPOsAsync(long vendId, string accNo);
        Task<MakerMasterPoModel> GetMasterPoDetailsAsync(string masterPoNo, bool isVendor);
        Task<string> SaveLoanAsync(MakerLongTermLoanDto dto);
    }
}
