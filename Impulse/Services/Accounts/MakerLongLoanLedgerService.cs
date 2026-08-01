using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public class MakerLongLoanLedgerService : IMakerLongLoanLedgerService
    {
        private readonly IMakerLongLoanLedgerDataAccess _dataAccess;

        public MakerLongLoanLedgerService(IMakerLongLoanLedgerDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<MakerAccountModel>> GetMakerAccountsAsync()
        {
            return _dataAccess.GetMakerAccountsAsync();
        }

        public async Task<MakerLongLoanLedgerResultModel> GetLedgerDataAsync(string accNo, string makerName, DateTime fromDate, DateTime toDate)
        {
            var result = new MakerLongLoanLedgerResultModel
            {
                AccNo = accNo,
                MakerName = makerName,
                FromDate = fromDate,
                ToDate = toDate
            };

            if (string.IsNullOrWhiteSpace(accNo))
            {
                return result;
            }

            decimal openingBal = await _dataAccess.GetOpeningBalanceAsync(accNo, fromDate);
            result.OpeningBalance = openingBal;

            var entries = await _dataAccess.GetLedgerEntriesAsync(accNo, fromDate, toDate);

            decimal runningBal = openingBal;
            decimal totalAdvances = 0m;
            decimal totalDeductions = 0m;

            foreach (var entry in entries)
            {
                if (entry.Amount >= 0)
                {
                    totalAdvances += entry.Amount;
                }
                else
                {
                    totalDeductions += Math.Abs(entry.Amount);
                }

                runningBal += entry.Amount;
                entry.RunningBalance = runningBal;
            }

            result.TotalAdvances = totalAdvances;
            result.TotalDeductions = totalDeductions;
            result.Entries = entries;

            return result;
        }
    }
}
