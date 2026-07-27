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
    public class FinishTransactionDataAccess : IFinishTransactionDataAccess
    {
        private readonly string _connectionString;

        public FinishTransactionDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<FinishTransactionModel>> GetTransactionsAsync(FinishTransactionFilterModel filter)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var sb = new StringBuilder();
                sb.Append(@"
                    SELECT EntryNo, DT, ISNULL(OrderNo, '') AS OrderNo, ISNULL(RcvdBy, '') AS RcvdBy, 
                           ISNULL(RcvdFrom, '') AS RcvdFrom, ItemID, ISNULL(ItemName, '') AS ItemName, 
                           ISNULL(Unit, '') AS Unit, Qty, ISNULL(Location, '') AS Location, 
                           ChildEntryID, EntryType, Shelf_RefID, Store_RefID, ISNULL(StoreName, '') AS StoreName, 
                           ISNULL(LotNo_Manual, '') AS LotNo_Manual, ISNULL(Remarks, '') AS Remarks
                    FROM VStockTransactions
                    WHERE CAST(DT AS DATE) BETWEEN @FromDate AND @ToDate");

                if (!filter.IncludeIssuance)
                {
                    sb.Append(" AND EntryType <> 0");
                }
                if (!filter.IncludeReceiving)
                {
                    sb.Append(" AND EntryType <> 1");
                }
                if (filter.StoreRefId > 0)
                {
                    sb.Append(" AND Store_RefID = @StoreRefId");
                }
                if (!string.IsNullOrWhiteSpace(filter.ItemID) && filter.ItemID != "0")
                {
                    sb.Append(" AND ItemID = @ItemID");
                }

                sb.Append(" ORDER BY DT DESC, EntryNo DESC");

                var results = await db.QueryAsync<FinishTransactionModel>(sb.ToString(), new
                {
                    FromDate = filter.FromDate.Date,
                    ToDate = filter.ToDate.Date,
                    StoreRefId = filter.StoreRefId,
                    ItemID = filter.ItemID
                });

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

        public async Task<List<ItemViewModel>> GetArticlesAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ItemID, '{' + ItemID + '} ' + ItemName AS ItemName, ISNULL(Unit, '') AS Unit
                    FROM Items
                    WHERE ISNULL(InActive, 0) = 0
                    ORDER BY ItemName";

                var results = await db.QueryAsync<ItemViewModel>(sql);
                return results.AsList();
            }
        }
    }
}
