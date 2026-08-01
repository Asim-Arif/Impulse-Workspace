using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Accounts
{
    public class MakerLongLoanLedgerDataAccess : IMakerLongLoanLedgerDataAccess
    {
        private readonly IConfiguration _config;

        public MakerLongLoanLedgerDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection");

        public async Task<List<MakerAccountModel>> GetMakerAccountsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VendID, AccNo, ISNULL(MakerNo, '') AS MakerNo, AccTitle AS MakerName, AccTitle
                           FROM VVendersAndMakers 
                           WHERE AccNo IS NOT NULL AND AccNo <> '' 
                           ORDER BY AccTitle";

            var list = (await db.QueryAsync<MakerAccountModel>(sql)).ToList();
            return list;
        }

        public async Task<decimal> GetOpeningBalanceAsync(string accNo, DateTime fromDate)
        {
            if (string.IsNullOrWhiteSpace(accNo)) return 0m;

            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ISNULL(SUM(Amount), 0) FROM VMakerAdvancesLedger WHERE AccNo = @AccNo AND DT < @FromDate";

            return await db.ExecuteScalarAsync<decimal>(sql, new { AccNo = accNo, FromDate = fromDate });
        }

        public async Task<List<MakerLongLoanLedgerEntryModel>> GetLedgerEntriesAsync(string accNo, DateTime fromDate, DateTime toDate)
        {
            if (string.IsNullOrWhiteSpace(accNo)) return new List<MakerLongLoanLedgerEntryModel>();

            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EntryID, AccNo, DT, Description, Remarks, Amount, DAmount, AccVoucherNo
                           FROM VMakerAdvancesLedger 
                           WHERE AccNo = @AccNo AND DT BETWEEN @FromDate AND @ToDate 
                           ORDER BY DT, Amount DESC";

            var list = (await db.QueryAsync<MakerLongLoanLedgerEntryModel>(sql, new { AccNo = accNo, FromDate = fromDate, ToDate = toDate })).ToList();
            return list;
        }
    }
}
