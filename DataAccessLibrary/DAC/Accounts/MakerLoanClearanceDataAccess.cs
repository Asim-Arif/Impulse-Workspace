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
    public class MakerLoanClearanceDataAccess : IMakerLoanClearanceDataAccess
    {
        private readonly IConfiguration _config;

        public MakerLoanClearanceDataAccess(IConfiguration config)
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
                           ORDER BY AccNo";

            var list = (await db.QueryAsync<MakerAccountModel>(sql)).ToList();
            return list;
        }

        public async Task<MakerLoanBalanceModel> GetMakerBalancesAsync(string accNo)
        {
            var model = new MakerLoanBalanceModel { AccNo = accNo };
            if (string.IsNullOrWhiteSpace(accNo)) return model;

            using IDbConnection db = new SqlConnection(ConnectionString);

            string ltSql = @"SELECT TotalAmountTaken, TotalAmountCleared, ActualDeductionAmount FROM VMakerAdvancesDeduction WHERE AccNo = @AccNo";
            var ltData = await db.QueryFirstOrDefaultAsync(ltSql, new { AccNo = accNo });

            if (ltData != null)
            {
                decimal totalTaken = Convert.ToDecimal(ltData.TotalAmountTaken ?? 0);
                decimal totalCleared = Convert.ToDecimal(ltData.TotalAmountCleared ?? 0);
                model.LTBalance = totalTaken - totalCleared;
                model.LTDeductionAmount = Convert.ToDecimal(ltData.ActualDeductionAmount ?? 0);
            }

            string stSql = @"SELECT Amount, AmountCleared FROM VMakerAdvancesShortDeduction WHERE AccNo = @AccNo";
            var stData = await db.QueryFirstOrDefaultAsync(stSql, new { AccNo = accNo });

            if (stData != null)
            {
                decimal amount = Convert.ToDecimal(stData.Amount ?? 0);
                decimal amountCleared = Convert.ToDecimal(stData.AmountCleared ?? 0);
                model.STBalance = amount - amountCleared;
            }

            return model;
        }

        public async Task SaveClearanceAsync(MakerLoanClearanceDto dto, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State == ConnectionState.Closed) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                long stEntryId = 0;
                long ltEntryId = 0;

                if (dto.STToClear > 0)
                {
                    string insertStSql = @"INSERT INTO MakerPostedBills_STDeductions(MPB_RefID, Amount, AccNo) VALUES(0, @Amount, @AccNo);
                                           SELECT CAST(SCOPE_IDENTITY() as bigint);";
                    stEntryId = await db.ExecuteScalarAsync<long>(insertStSql, new { Amount = dto.STToClear, AccNo = dto.AccNo }, trans);
                }

                if (dto.LTToClear > 0)
                {
                    string insertLtSql = @"INSERT INTO MakerAmtCleared(AccNo, DT, AmtClrd) VALUES(@AccNo, @DT, @AmtClrd);
                                           SELECT CAST(SCOPE_IDENTITY() as bigint);";
                    ltEntryId = await db.ExecuteScalarAsync<long>(insertLtSql, new { AccNo = dto.AccNo, DT = dto.ClearanceDate, AmtClrd = dto.LTToClear }, trans);
                }

                string insertAuditSql = @"INSERT INTO MakerLoanClearance_Manual(DT, AccNo, MPB_STD_EntryID, MAC_EntryID, UserName, MachineName, Remarks)
                                          VALUES(@DT, @AccNo, @MPB_STD_EntryID, @MAC_EntryID, @UserName, @MachineName, @Remarks)";

                await db.ExecuteAsync(insertAuditSql, new
                {
                    DT = dto.ClearanceDate,
                    AccNo = dto.AccNo,
                    MPB_STD_EntryID = stEntryId,
                    MAC_EntryID = ltEntryId,
                    UserName = userName,
                    MachineName = machineName,
                    Remarks = dto.Remarks
                }, trans);

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task UpdateLTDeductionAmountAsync(string accNo, decimal newDAmount)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"UPDATE MakerAdvances 
                           SET DAmount = @DAmount 
                           WHERE EntryID IN (SELECT MAX(EntryID) FROM MakerAdvances WHERE AccNo = @AccNo)";

            await db.ExecuteAsync(sql, new { AccNo = accNo, DAmount = newDAmount });
        }
    }
}
