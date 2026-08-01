using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Accounts
{
    public interface IMakerLongTermLoanDataAccess
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<decimal> GetMakerLongTermBalanceAsync(string accNo);
        Task<List<MakerMasterPoModel>> GetMakerMasterPOsAsync(long vendId, string accNo);
        Task<MakerMasterPoModel> GetMasterPoDetailsAsync(string masterPoNo, bool isVendor);
        Task SaveMakerLongTermLoanAsync(MakerLongTermLoanDto dto, string vchrNo);
    }
}
