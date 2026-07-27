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
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class FinishItemLedgerDataAccess : IFinishItemLedgerDataAccess
    {
        private readonly string _connectionString;

        public FinishItemLedgerDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<ItemHeaderInfoModel?> GetItemHeaderInfoAsync(string itemId)
        {
            if (string.IsNullOrWhiteSpace(itemId)) return null;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string itemSql = @"
                    SELECT ItemID, ItemName, ISNULL(Unit, '') AS Unit, 
                           ISNULL(ItemSize, '') AS ItemSize, ISNULL(SizeUnit, '') AS SizeUnit, 
                           ISNULL(ItemColor, '') AS ItemColor
                    FROM VItems
                    WHERE ItemID = @ItemID";

                var info = await db.QueryFirstOrDefaultAsync<ItemHeaderInfoModel>(itemSql, new { ItemID = itemId });
                if (info != null)
                {
                    string stockSql = @"
                        SELECT ISNULL(SUM(NetQty), 0) 
                        FROM VItemsWithShelfWiseStock 
                        WHERE ItemID = @ItemID";
                    info.TotalStock = await db.ExecuteScalarAsync<double>(stockSql, new { ItemID = itemId });
                }
                return info;
            }
        }

        public async Task<List<FinishItemLedgerRowModel>> GetLedgerAsync(FinishItemLedgerFilterModel filter)
        {
            if (string.IsNullOrWhiteSpace(filter?.ItemID))
                return new List<FinishItemLedgerRowModel>();

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var p = new DynamicParameters();
                p.Add("@ItemID", filter.ItemID);
                p.Add("@DTFrom", filter.FromDate.Date);
                p.Add("@DTTo", filter.ToDate.Date);

                var results = (await db.QueryAsync<FinishItemLedgerRowModel>(
                    "FinishItemLedger_SP",
                    p,
                    commandType: CommandType.StoredProcedure
                )).AsList();

                return results;
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
