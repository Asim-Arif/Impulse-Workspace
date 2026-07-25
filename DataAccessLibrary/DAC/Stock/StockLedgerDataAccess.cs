using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class StockLedgerDataAccess : IStockLedgerDataAccess
    {
        private readonly string _connectionString;

        public StockLedgerDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<RMViewModel>> GetMaterialsAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT RMID1, '{' + RMID1 + '} ' + RMName AS RMName
                    FROM RM 
                    WHERE ISNULL(InActive, 0) = 0
                    ORDER BY RMID1";
                var results = await db.QueryAsync<RMViewModel>(sql);
                return results.AsList();
            }
        }

        public async Task<double> GetOpeningBalanceAsync(string materialId, DateTime fromDate)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                double openingBalance = 0;

                // 1. Get Sum of Quantity from VStockLedger where Date < fromDate
                string sumSql = @"
                    SELECT ISNULL(SUM(Quantity), 0) 
                    FROM VStockLedger 
                    WHERE MaterialID = @MaterialId AND TransactionDate < @FromDate";
                double sumQty = await db.QuerySingleOrDefaultAsync<double>(sumSql, new { MaterialId = materialId, FromDate = fromDate });

                // 2. Get OpeningStock from RM table
                string rmSql = @"
                    SELECT ISNULL(OpeningStock, 0) 
                    FROM RM 
                    WHERE RMID1 = @MaterialId";
                double openingStock = await db.QuerySingleOrDefaultAsync<double>(rmSql, new { MaterialId = materialId });

                openingBalance = sumQty + openingStock;
                return Math.Round(openingBalance, 2);
            }
        }

        public async Task<List<StockLedgerTransactionModel>> GetTransactionsAsync(string materialId, DateTime fromDate, DateTime toDate)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT TransactionID, TransactionDate, MaterialID, Description, Reference, 
                           Quantity, TransactionType, Rate, UserName, RecieptID, IssuanceType, 
                           POIssuance, Sheets
                    FROM VStockLedger
                    WHERE MaterialID = @MaterialId 
                      AND TransactionDate >= @FromDate 
                      AND TransactionDate <= @ToDate
                    ORDER BY TransactionDate, TransactionID";

                var results = await db.QueryAsync<StockLedgerTransactionModel>(sql, new 
                { 
                    MaterialId = materialId, 
                    FromDate = fromDate.Date, 
                    ToDate = toDate.Date 
                });

                return results.AsList();
            }
        }
    }
}
