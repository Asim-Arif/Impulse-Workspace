using DataAccessLibrary.Models.ViewModels.Accounts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Accounts
{
    public interface IMakerLongLoanLedgerDataAccess
    {
        Task<List<MakerAccountModel>> GetMakerAccountsAsync();
        Task<decimal> GetOpeningBalanceAsync(string accNo, DateTime fromDate);
        Task<List<MakerLongLoanLedgerEntryModel>> GetLedgerEntriesAsync(string accNo, DateTime fromDate, DateTime toDate);
    }
}
