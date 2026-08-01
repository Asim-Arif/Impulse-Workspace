using DataAccessLibrary.Models.ViewModels.Accounts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Accounts
{
    public interface IMakerShortLoanLedgerDataAccess
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<decimal> GetOpeningBalanceAsync(string accNo, DateTime fromDate);
        Task<List<MakerShortLoanLedgerEntryModel>> GetLedgerEntriesAsync(string accNo, DateTime fromDate, DateTime toDate);
    }
}
