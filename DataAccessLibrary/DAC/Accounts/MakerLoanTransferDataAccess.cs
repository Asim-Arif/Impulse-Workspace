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
    public class MakerLoanTransferDataAccess : IMakerLoanTransferDataAccess
    {
        private readonly IConfiguration _config;

        public MakerLoanTransferDataAccess(IConfiguration config)
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

        public async Task SaveTransferAsync(MakerLoanTransferDto dto, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State == ConnectionState.Closed) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                string insertMasterSql = @"INSERT INTO MakerLoanTransfer(AccNo_From, AccNo_To, ST_Amt, LT_Amt, UserName, MachineName, Description, DT, DTEntry)
                                            VALUES(@AccNo, @AccNo, @ST_Amt, @LT_Amt, @UserName, @MachineName, @Description, @DT, GETDATE());
                                            SELECT CAST(SCOPE_IDENTITY() as bigint);";

                long mltRefId = await db.ExecuteScalarAsync<long>(insertMasterSql, new
                {
                    AccNo = dto.AccNo,
                    ST_Amt = dto.STToClear,
                    LT_Amt = dto.LTToClear,
                    UserName = userName,
                    MachineName = machineName,
                    Description = dto.Remarks,
                    DT = dto.TransferDate
                }, trans);

                // ── S.T to L.T Transfer ──────────────────────────────
                if (dto.STToClear > 0)
                {
                    string insertStDedSql = @"INSERT INTO MakerPostedBills_STDeductions(MPB_RefID, Amount, AccNo, MLT_RefID)
                                              VALUES(0, @Amount, @AccNo, @MLT_RefID)";
                    await db.ExecuteAsync(insertStDedSql, new { Amount = dto.STToClear, AccNo = dto.AccNo, MLT_RefID = mltRefId }, trans);

                    string insertLtAdvSql = @"INSERT INTO MakerAdvances(AccNo, DT, Type, Description, Amount, DAmount, AmountCleared, AccVoucherNo, MLT_RefID)
                                              VALUES(@AccNo, @DT, 1, @Description, @Amount, @DAmount, 0, '', @MLT_RefID)";
                    await db.ExecuteAsync(insertLtAdvSql, new
                    {
                        AccNo = dto.AccNo,
                        DT = dto.TransferDate,
                        Description = dto.Remarks,
                        Amount = dto.STToClear,
                        DAmount = dto.NewLTDeductionAmount,
                        MLT_RefID = mltRefId
                    }, trans);
                }

                // ── L.T to S.T Transfer ──────────────────────────────
                if (dto.LTToClear > 0)
                {
                    string insertLtClrdSql = @"INSERT INTO MakerAmtCleared(AccNo, DT, AmtClrd, MLT_RefID)
                                               VALUES(@AccNo, @DT, @AmtClrd, @MLT_RefID)";
                    await db.ExecuteAsync(insertLtClrdSql, new { AccNo = dto.AccNo, DT = dto.TransferDate, AmtClrd = dto.LTToClear, MLT_RefID = mltRefId }, trans);

                    string insertStAdvSql = @"INSERT INTO MakerAdvancesShort(AccNo, DT, Description, Amount, AmountCleared, AccVoucherNo, MLT_RefID)
                                              VALUES(@AccNo, @DT, @Description, @Amount, 0, '', @MLT_RefID)";
                    await db.ExecuteAsync(insertStAdvSql, new
                    {
                        AccNo = dto.AccNo,
                        DT = dto.TransferDate,
                        Description = dto.Remarks,
                        Amount = dto.LTToClear,
                        MLT_RefID = mltRefId
                    }, trans);
                }

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }
    }
}
