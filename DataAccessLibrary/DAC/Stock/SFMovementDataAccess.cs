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
    public class SFMovementDataAccess : ISFMovementDataAccess
    {
        private readonly string _connectionString;

        public SFMovementDataAccess(IConfiguration config)
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
                    ORDER BY ItemID";
                var results = await db.QueryAsync<ItemViewModel>(sql);
                return results.AsList();
            }
        }

        public async Task<List<ItemProcessViewModel>> GetProcessesByArticleAsync(string itemId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ProcessID, Description, ISNULL(ItemSNo, 0) AS ItemSNo
                    FROM VItemProcesses
                    WHERE IsExist = @ItemId
                    ORDER BY ItemSNo";
                var results = await db.QueryAsync<ItemProcessViewModel>(sql, new { ItemId = itemId });
                return results.AsList();
            }
        }

        public async Task<List<SFTargetStoreModel>> GetFromStoresAsync(string itemId, int processId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DISTINCT Store_RefID AS EntryID, StoreName
                    FROM VStockOrderOpening
                    WHERE ItemID = @ItemId 
                      AND ProcessID = @ProcessId
                      AND (Qty - ISNULL(IssdQty, 0)) > 0
                    ORDER BY StoreName";
                var results = await db.QueryAsync<SFTargetStoreModel>(sql, new { ItemId = itemId, ProcessId = processId });
                return results.AsList();
            }
        }

        public async Task<List<SFFromLocationModel>> GetFromLocationsAsync(string itemId, int processId, int storeRefId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, ItemID, ProcessID, Store_RefID, Shelf_RefID,
                           ISNULL(RackNo, '') AS RackNo, ISNULL(ShelfNo, '') AS ShelfNo,
                           ISNULL(StoreName, '') AS StoreName, ISNULL(LotNo_Manual, '') AS LotNo_Manual,
                           (Qty - ISNULL(IssdQty, 0)) AS AvailableQty, ISNULL(Remarks, '') AS Remarks
                    FROM VStockOrderOpening
                    WHERE ItemID = @ItemId 
                      AND ProcessID = @ProcessId 
                      AND Store_RefID = @StoreRefId
                      AND (Qty - ISNULL(IssdQty, 0)) > 0
                    ORDER BY EntryID";
                var results = await db.QueryAsync<SFFromLocationModel>(sql, new { ItemId = itemId, ProcessId = processId, StoreRefId = storeRefId });
                return results.AsList();
            }
        }

        public async Task<List<SFTargetStoreModel>> GetTargetStoresAsync(string userName)
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
                var results = await db.QueryAsync<SFTargetStoreModel>(sql, new { UserName = userName });
                if (!results.Any())
                {
                    results = await db.QueryAsync<SFTargetStoreModel>("SELECT EntryID, StoreName FROM Stores ORDER BY StoreName");
                }
                return results.AsList();
            }
        }

        public async Task<List<SFTargetRackModel>> GetTargetRacksAsync(int storeRefId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DISTINCT ISNULL(RackNo, '') AS RackNo
                    FROM VStoreShelfs
                    WHERE Store_RefID = @StoreRefId
                    ORDER BY RackNo";
                var results = await db.QueryAsync<SFTargetRackModel>(sql, new { StoreRefId = storeRefId });
                return results.AsList();
            }
        }

        public async Task<List<SFTargetShelfModel>> GetTargetShelvesAsync(int storeRefId, string rackNo)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, Store_RefID, ISNULL(RackNo, '') AS RackNo, ISNULL(ShelfNo, '') AS ShelfNo
                    FROM VStoreShelfs
                    WHERE Store_RefID = @StoreRefId AND ISNULL(RackNo, '') = @RackNo
                    ORDER BY ShelfNo";
                var results = await db.QueryAsync<SFTargetShelfModel>(sql, new { StoreRefId = storeRefId, RackNo = rackNo });
                return results.AsList();
            }
        }

        public async Task<int> GetTargetLocationQtyAsync(string itemId, int processId, int shelfRefId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ISNULL(SUM(Qty - ISNULL(IssdQty, 0)), 0)
                    FROM VStockOrderOpening
                    WHERE Shelf_RefID = @ShelfRefId AND ItemID = @ItemId AND ProcessID = @ProcessId";
                int qty = await db.ExecuteScalarAsync<int>(sql, new { ShelfRefId = shelfRefId, ItemId = itemId, ProcessId = processId });
                return qty;
            }
        }

        public async Task<bool> TransferLocationAsync(SFMovementRequest request)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    UPDATE StockOrderOpening 
                    SET Location = @TargetLocationDisplay, Shelf_RefID = @TargetShelfID
                    WHERE EntryID = @SOO_EntryID";
                int rows = await db.ExecuteAsync(sql, request);
                return rows > 0;
            }
        }
    }
}
