using DataAccessLibrary.Models.ViewModels.Accounts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public interface IMakerShortLoanLedgerService
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<MakerShortLoanLedgerResultModel> GetLedgerDataAsync(string accNo, string makerName, DateTime fromDate, DateTime toDate);
    }
}
