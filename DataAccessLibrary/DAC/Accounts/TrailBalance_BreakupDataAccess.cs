using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Linq.Expressions;
using System.Reflection.PortableExecutable;
using System.Threading.Tasks;

public class TrailBalance_BreakupDataAccess : ITrailBalance_BreakupDataAccess
{
    private readonly string _connectionString;
    private readonly ILogger<FinancialStatementDataAccess> _logger;

    public TrailBalance_BreakupDataAccess(IConfiguration config)
    {
        
        _connectionString = config.GetConnectionString("DefaultConnection");
    }

    public async Task<List<MainHeadsModel>> GetMainHeadsList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY SortNo ASC";
                string sql = @" SELECT EntryID,GroupName,SortNo,SequenceNo,Description FROM GroupsForFinancialReport " + cond;
                var entrieslist = await db.QueryAsync<MainHeadsModel>(sql, new { });
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching: {ex.Message}");
            throw;
        }
    }

    public async Task<List<TrailBalance_BreakupModel>> GetAccountsList(DateTime fromDate, DateTime toDate, bool isSummary, string accountCode = "")
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            var para = new DynamicParameters();
            para.Add("@FromDate", fromDate);
            para.Add("@Summary", isSummary ? 1 : 0);
            if (!string.IsNullOrEmpty(accountCode)) para.Add("@Code", accountCode);

            string spName = "TBBase_Breakup";

            var accounts = await db.QueryAsync<TrailBalance_BreakupModel>(spName, para, commandType: CommandType.StoredProcedure, commandTimeout : 240);
            return accounts.ToList();
        }
    }

    public async Task<TrailBalance_BreakupModel> CalculateAccountBalance_Query(string accNo, DateTime fromDate, DateTime toDate, bool isParent)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            var result = new TrailBalance_BreakupModel();

            string accListSql = "";
            int timeout = 240;

            if (isParent)
            {

                string childQuery = "SELECT AccNo FROM Accounts WHERE SubAccOf = @AccNo";
                var children = await db.QueryAsync<string>(childQuery, new { AccNo = accNo },commandTimeout : timeout);

                var accList = children.ToList();
                accList.Add(accNo); 
                accListSql = string.Join(",", accList.Select(x => $"'{x}'"));
            }
            else
            {
                accListSql = $"'{accNo}'";
            }

            string voucherSql = $@"
                SELECT Sum(Debit) as Debit, Sum(Credit) as Credit
                FROM Vouchers 
                WHERE AccNo IN ({accListSql}) 
                AND (Credit IS NOT NULL AND Debit IS NOT NULL)
                AND (VDate BETWEEN @FromDate AND @ToDate) 
                AND LEFT(VchrNo, 2) <> 'CL'";

            var transTotals = await db.QueryFirstOrDefaultAsync(voucherSql, new { FromDate = fromDate, ToDate = toDate }, commandTimeout: timeout);
            result.Debit = transTotals?.Dr ?? 0;
            result.Credit = transTotals?.Cr ?? 0;

            string adjSql = $@"
                SELECT Sum(Debit) as AdjDr, Sum(Credit) as AdjCr
                FROM VVouchers_Detail_Adj 
                WHERE AccNo IN ({accListSql}) 
                AND (Credit IS NOT NULL AND Debit IS NOT NULL)
                AND (DT BETWEEN @FromDate AND @ToDate)";

            var adjTotals = await db.QueryFirstOrDefaultAsync(adjSql, new { FromDate = fromDate, ToDate = toDate }, commandTimeout: timeout);
            result.AdjDr = adjTotals?.AdjDr ?? 0;
            result.AdjCr = adjTotals?.AdjCr ?? 0;

            if (isParent)
            {
                string complexBalSql = $@"
                    SELECT 
                        SUM(ISNULL(TBase.TheBalance, A.openbal)) AS OpeningBalance,
                        SUM(CASE WHEN ISNULL(TBase.TheBalance, A.openbal)>=0 THEN ISNULL(TBase.TheBalance, A.openbal) ELSE 0 END) AS OpeningBalanceDr,
                        SUM(CASE WHEN ISNULL(TBase.TheBalance, A.openbal)<0 THEN ABS(ISNULL(TBase.TheBalance, A.openbal)) ELSE 0 END) AS OpeningBalanceCr,
                        SUM(CASE WHEN ISNULL(TBaseBal.TheBalance, A.OpenBal)>=0 THEN ISNULL(TBaseBal.TheBalance, A.OpenBal) ELSE 0 END) AS LastBalanceBalanceDr,
                        SUM(CASE WHEN ISNULL(TBaseBal.TheBalance, A.OpenBal)<0 THEN ABS(ISNULL(TBaseBal.TheBalance, A.OpenBal)) ELSE 0 END) AS LastBalanceBalanceCr
                    FROM dbo.Accounts A
                    LEFT OUTER JOIN 
                        (SELECT Accno, balance AS TheBalance 
                         FROM dbo.Vouchers 
                         INNER JOIN (SELECT MAX(Sno) AS MAXSNo FROM Vouchers WHERE VDATE < @FromDate GROUP BY Accno) TMaxSno1 
                         ON Vouchers.SNo=TMaxSno1.MAXSNo 
                         WHERE Vouchers.AccNo IN ({accListSql})) TBase 
                    ON A.AccNo = TBase.Accno 
                    LEFT OUTER JOIN 
                        (SELECT Accno, balance AS TheBalance 
                         FROM dbo.Vouchers 
                         INNER JOIN (
                            SELECT MAX(Sno) AS MAXSNo FROM Vouchers 
                            LEFT JOIN (SELECT AccNo, COUNT(*) AS CLEntries FROM Vouchers WHERE LEFT(VchrNo,2)='CL' AND VDATE BETWEEN @FromDate AND @ToDate GROUP BY Vouchers.Accno) TCl 
                            ON Vouchers.AccNo=TCl.AccNo 
                            WHERE VDATE <= @ToDate 
                            AND LEFT(VchrNo,2) <> (CASE WHEN ISNULL(TCl.CLEntries,0)>0 THEN 'CL' ELSE '' END) 
                            GROUP BY Vouchers.Accno
                         ) TMaxSNo2 
                         ON Vouchers.SNo=TMaxSNo2.MAXSNo 
                         WHERE Vouchers.AccNo IN ({accListSql})) TBaseBal 
                    ON A.AccNo = TBaseBal.Accno 
                    WHERE A.AccNo IN ({accListSql})";

                //string acclistforcond = accListSql.Replace("'", "");
                //string complexBalSql = @" 
                //        SELECT OpeningBalance, OpeningBalanceDr, OpeningBalanceCr,
                //        LastBalanceBalanceDr, LastBalanceBalanceCr
                //        FROM dbo.FGetTrailBalance_Data(@FromDate, @ToDate, @AccList)";

                var balData = await db.QueryFirstOrDefaultAsync(complexBalSql, new
                {
                    FromDate = fromDate,
                    ToDate = toDate,
                    //AccList = acclistforcond
                } , commandTimeout: timeout);

                if (balData != null)
                {
                    result.OpeningDr = (double)(balData.OpeningBalanceDr ?? 0);
                    result.OpeningCr = (double)(balData.OpeningBalanceCr ?? 0);
                    result.ClosingDr = (double)(balData.LastBalanceBalanceDr ?? 0);
                    result.ClosingCr = (double)(balData.LastBalanceBalanceCr ?? 0);
                }
            }
            else
            {

                string simpleOpSql = "SELECT OpenBal FROM Accounts WHERE AccNo = @AccNo";
                double dbOpBal = await db.QueryFirstOrDefaultAsync<double>(simpleOpSql, new { AccNo = accNo } , commandTimeout: timeout);

                double netOp = dbOpBal;
                if (netOp >= 0) result.OpeningDr = netOp; else result.OpeningCr = Math.Abs(netOp);

                double netClose = netOp + result.Debit - result.Credit;
                if (netClose >= 0) result.ClosingDr = netClose; else result.ClosingCr = Math.Abs(netClose);
            }

            return result;
        }
    }

    public async Task<List<TrailBalance_BreakupModel>> CalculateAccountBalances(string code, DateTime fromDate, DateTime toDate)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            var p = new DynamicParameters();
            p.Add("@FromDate", fromDate);
            p.Add("@ToDate", toDate);
            p.Add("@Code", string.IsNullOrEmpty(code) ? null : code);

            var result = await db.QueryAsync<TrailBalance_BreakupModel>(
                "TBBase_Accounts_Balance_SP",
                p,
                commandType: CommandType.StoredProcedure,
                commandTimeout: 240
            );
            return result.ToList();
        }
    }

    public async Task SaveLoadedDataToTempTable(List<TrailBalance_BreakupModel> modeldata, bool bWith_Opening_Closing)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            string sql = "";
            if (bWith_Opening_Closing)
            {
                await db.ExecuteAsync("DELETE FROM TempTrialBal");
                sql = @"INSERT INTO TempTrialBal 
                            (AccNo,AccTitle,OpenBal,Debit,Credit,LastBal,Bold) 
                            VALUES 
                            (@AccNo, @AccTitle, (@OpeningDr-@OpeningCr), @Debit, @Credit, (@ClosingDr-@ClosingCr), @Parent)";
            }
            else
            {
                await db.ExecuteAsync("DELETE FROM TempTrialBalanceNew");
                sql = @"INSERT INTO TempTrialBalanceNew 
                            (AccNo, AccTitle, OpenBalDr, OpenBalCr, Debit, Credit, LastBalDr, LastBalCr, 
                            Bold, Debit_Adj, Credit_Adj, LastBalDr_Adj, LastBalCr_Adj) 
                            VALUES 
                            (@AccNo, @AccTitle, @OpeningDr, @OpeningCr, @Debit, @Credit, @ClosingDr, 
                             @ClosingCr, @Parent, @AdjDr, @AdjCr, @FinalClosingDr, @FinalClosingCr )";
            }
            await db.ExecuteAsync(sql, modeldata);
        }
    }

}
