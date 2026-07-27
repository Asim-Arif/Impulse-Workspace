using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class SFTransactionDataAccess : ISFTransactionDataAccess
    {
        private readonly string _connectionString;

        public SFTransactionDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<ItemViewModel>> GetArticlesAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ItemID, '{' + ItemID + '} ' + ItemName AS ItemName
                    FROM Items 
                    WHERE ISNULL(InActive, 0) = 0 
                    ORDER BY ItemName";
                var results = await db.QueryAsync<ItemViewModel>(sql);
                return results.AsList();
            }
        }

        public async Task<List<ProcessLookupModel>> GetProcessesAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ProcessID, Description, ISNULL(Code, '') AS Code
                    FROM Processes
                    ORDER BY ISNULL(SNo, 0)";
                var results = await db.QueryAsync<ProcessLookupModel>(sql);
                return results.AsList();
            }
        }

        public async Task<List<StoreLookupModel>> GetStoresAsync(string userName)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DISTINCT s.EntryID, s.StoreName
                    FROM Stores s
                    LEFT JOIN Users_Stores us ON s.EntryID = us.Store_RefID
                    LEFT JOIN Users u ON us.UserID = u.UserID
                    WHERE u.UserName = @UserName OR NOT EXISTS (SELECT 1 FROM Users_Stores)
                    ORDER BY s.StoreName";
                var results = await db.QueryAsync<StoreLookupModel>(sql, new { UserName = userName });
                if (!results.Any())
                {
                    results = await db.QueryAsync<StoreLookupModel>("SELECT EntryID, StoreName FROM Stores ORDER BY StoreName");
                }
                return results.AsList();
            }
        }

        public async Task<List<SFTransactionModel>> GetTransactionsAsync(SFTransactionFilterModel filter)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var sb = new StringBuilder();
                sb.AppendLine(@"
                    SELECT EntryID, ISNULL(UserName, '') AS UserName, DT, ISNULL(ItemCode, '') AS ItemCode,
                           ISNULL(ItemName, '') AS ItemName, ProcessID, ISNULL(Description, '') AS Description,
                           ISNULL(QtyIssd, 0) AS QtyIssd, ISNULL(Issuance, 0) AS Issuance,
                           ISNULL(Location, '') AS Location, ISNULL(StoreName, '') AS StoreName,
                           Store_RefID, ISNULL(LotNo, '') AS LotNo
                    FROM VSFStockTransaction
                    WHERE DT >= @FromDate AND DT <= @ToDate");

                if (!filter.IncludeIssuance)
                {
                    sb.AppendLine("  AND Issuance <> 1");
                }
                if (!filter.IncludeReceiving)
                {
                    sb.AppendLine("  AND Issuance <> 0");
                }
                if (filter.StoreRefId > 0)
                {
                    sb.AppendLine("  AND Store_RefID = @StoreRefId");
                }
                if (!string.IsNullOrEmpty(filter.ItemCode) && filter.ItemCode != "0")
                {
                    sb.AppendLine("  AND ItemCode = @ItemCode");
                }
                if (filter.ProcessId > 0)
                {
                    sb.AppendLine("  AND ProcessID = @ProcessId");
                }

                sb.AppendLine("ORDER BY DT DESC, EntryID DESC");

                var parameters = new
                {
                    FromDate = filter.FromDate.Date,
                    ToDate = filter.ToDate.Date.AddDays(1).AddTicks(-1),
                    StoreRefId = filter.StoreRefId,
                    ItemCode = filter.ItemCode,
                    ProcessId = filter.ProcessId
                };

                var results = await db.QueryAsync<SFTransactionModel>(sb.ToString(), parameters);
                return results.AsList();
            }
        }
    }
}
