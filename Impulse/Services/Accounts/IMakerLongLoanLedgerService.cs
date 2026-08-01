using DataAccessLibrary.Models.ViewModels.Accounts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public interface IMakerLongLoanLedgerService
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<MakerLongLoanLedgerResultModel> GetLedgerDataAsync(string accNo, string makerName, DateTime fromDate, DateTime toDate);
    }
}
