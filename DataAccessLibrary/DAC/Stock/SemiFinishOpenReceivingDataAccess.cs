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
    public class SemiFinishOpenReceivingDataAccess : ISemiFinishOpenReceivingDataAccess
    {
        private readonly string _connectionString;

        public SemiFinishOpenReceivingDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<ItemViewModel>> GetArticlesAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ItemID, '{ ' + ItemID + ' } ' + ItemName AS ItemName
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

        public async Task<List<StoreLookupModel>> GetUserStoresAsync(string userName)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // Fetch stores available to the user or all active stores if not restricted
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
                    // Fallback to all stores if mapping query returned empty
                    results = await db.QueryAsync<StoreLookupModel>("SELECT EntryID, StoreName FROM Stores ORDER BY StoreName");
                }
                return results.AsList();
            }
        }

        public async Task<List<StoreShelfLookupModel>> GetStoreShelvesAsync(int storeRefId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, Store_RefID, ISNULL(RackNo, '') AS RackNo, ISNULL(ShelfNo, '') AS ShelfNo, ISNULL(StoreName, '') AS StoreName
                    FROM VStoreShelfs
                    WHERE Store_RefID = @StoreRefId
                    ORDER BY RackNo, ShelfNo";
                var results = await db.QueryAsync<StoreShelfLookupModel>(sql, new { StoreRefId = storeRefId });
                return results.AsList();
            }
        }

        public async Task<List<SemiFinishOpeningModel>> GetOpeningsByArticleAsync(string itemId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, ItemID, ItemName, ProcessID, Description, Qty, ISNULL(IssdQty, 0) AS IssdQty,
                           Location, StoreName, Store_RefID, Shelf_RefID, Remarks, DTEntry,
                           ISNULL(LotNo_Manual, '') AS LotNo_Manual,
                           ISNULL(BatchNo_Manual, '') AS BatchNo_Manual,
                           ISNULL(Mill_Certificate_No, '') AS Mill_Certificate_No,
                           UserName, MachineName
                    FROM VStockOrderOpening
                    WHERE ItemID = @ItemId AND Qty > ISNULL(IssdQty, 0)
                    ORDER BY EntryID DESC";
                var results = await db.QueryAsync<SemiFinishOpeningModel>(sql, new { ItemId = itemId });
                return results.AsList();
            }
        }

        public async Task<string> GetNextLotNoAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string year = DateTime.Today.ToString("yy");

                string sql1 = "SELECT ISNULL(MAX(CAST(SUBSTRING(LotNo, 3, CASE WHEN CHARINDEX('-', LotNo)>0 THEN CHARINDEX('-', LotNo)-3 ELSE LEN(LotNo) END) AS INT)), 0) FROM Lots_List WHERE LEFT(LotNo, 2) = @Year AND LEN(LotNo) = 7 AND ISNUMERIC(LotNo) = 1";
                string sql2 = "SELECT ISNULL(MAX(CAST(SUBSTRING(LotNo, 7, CASE WHEN CHARINDEX('-', LotNo)>0 THEN CHARINDEX('-', LotNo)-3 ELSE LEN(LotNo) END) AS INT)), 0) FROM VendRcvdDetail WHERE LEFT(LotNo, 2) = @Year AND LEN(LotNo) = 7 AND ISNUMERIC(LotNo) = 1";
                string sql3 = "SELECT ISNULL(MAX(CAST(SUBSTRING(LotNo, 7, CASE WHEN CHARINDEX('-', LotNo)>0 THEN CHARINDEX('-', LotNo)-3 ELSE LEN(LotNo) END) AS INT)), 0) FROM VendIssdDetail WHERE LEFT(LotNo, 2) = @Year AND LEN(LotNo) = 7 AND ISNUMERIC(LotNo) = 1";
                string sql4 = "SELECT ISNULL(MAX(CAST(SUBSTRING(LotNo_Manual, 7, CASE WHEN CHARINDEX('-', LotNo_Manual)>0 THEN CHARINDEX('-', LotNo_Manual)-3 ELSE LEN(LotNo_Manual) END) AS INT)), 0) FROM StockOrderOpening WHERE LEFT(LotNo_Manual, 2) = @Year AND LEN(LotNo_Manual) = 7 AND ISNUMERIC(LotNo_Manual) = 1";

                int max1 = await db.QuerySingleOrDefaultAsync<int>(sql1, new { Year = year });
                int max2 = await db.QuerySingleOrDefaultAsync<int>(sql2, new { Year = year });
                int max3 = await db.QuerySingleOrDefaultAsync<int>(sql3, new { Year = year });
                int max4 = await db.QuerySingleOrDefaultAsync<int>(sql4, new { Year = year });

                int maxVal = Math.Max(Math.Max(max1, max2), Math.Max(max3, max4));
                int nextVal = maxVal + 1;

                return $"{year}{nextVal:D5}";
            }
        }

        public async Task<bool> CheckBatchNoExistsAsync(string batchNo)
        {
            if (string.IsNullOrWhiteSpace(batchNo)) return false;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT COUNT(1) FROM StockOrderOpening WHERE BatchNo_Manual = @BatchNo";
                int count = await db.ExecuteScalarAsync<int>(sql, new { BatchNo = batchNo });
                return count > 0;
            }
        }

        public async Task<int> AddOpeningAsync(CreateSemiFinishOpeningRequest request)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        // 1. Insert StockOrderOpening
                        string insertSOOSql = @"
                            INSERT INTO StockOrderOpening 
                                (ItemID, ProcessID, Qty, Location, UserName, MachineName, Shelf_RefID, Remarks, LotNo_Manual, BatchNo_Manual, Mill_Certificate_No, DTEntry)
                            VALUES 
                                (@ItemID, @ProcessID, @Qty, @Location, @UserName, @MachineName, @Shelf_RefID, @Remarks, @LotNo_Manual, @BatchNo_Manual, @Mill_Certificate_No, GETDATE());
                            SELECT CAST(SCOPE_IDENTITY() as int);";

                        int sooEntryId = await db.ExecuteScalarAsync<int>(insertSOOSql, request, trans);

                        // 2. Insert Lots_List
                        string insertLotSql = @"
                            IF NOT EXISTS (SELECT 1 FROM Lots_List WHERE LotNo = @LotNo_Manual)
                            BEGIN
                                INSERT INTO Lots_List (LotNo, ItemID, Lot_Type, Reference_LotNo, Batch_No, Mill_Certificate_No)
                                VALUES (@LotNo_Manual, @ItemID, 1, '', @BatchNo_Manual, @Mill_Certificate_No);
                            END";
                        await db.ExecuteAsync(insertLotSql, request, trans);

                        // 3. Foreign Customer Boilerplate
                        string fcSql = @"
                            IF NOT EXISTS (SELECT 1 FROM ForeignCustomers WHERE Custcode='Stock' AND Country='PK')
                                INSERT INTO ForeignCustomers(Custcode, Country) VALUES('Stock', 'PK');

                            IF NOT EXISTS (SELECT 1 FROM FCustomerOrders WHERE OrderNo='Stock-Order')
                                INSERT INTO FCustomerOrders(CustCode, Country, OrderNo, DT) VALUES('Stock', 'PK', 'Stock-Order', GETDATE());

                            IF NOT EXISTS (SELECT 1 FROM FCustomerCatalog WHERE Custcode='Stock' AND Country='PK' AND CompItemID=@ItemID)
                                INSERT INTO FCustomerCatalog(Custcode, Country, ItemID, CompItemID) VALUES('Stock', 'PK', @ItemID, @ItemID);

                            IF NOT EXISTS (SELECT 1 FROM FOrderItems WHERE OrderNo='Stock-Order' AND CompItemCode=@ItemID)
                                INSERT INTO FOrderItems(OrderNo, ItemCode, CompItemCode, Qty, DeliveryDT, Stamps, Quality) VALUES('Stock-Order', @ItemID, @ItemID, @Qty, GETDATE(), '', '');
                            ELSE
                                UPDATE FOrderItems SET Qty = Qty + @Qty WHERE OrderNo='Stock-Order' AND ItemCode=@ItemID;";
                        await db.ExecuteAsync(fcSql, new { ItemID = request.ItemID, Qty = request.Qty }, trans);

                        // 4. VendReceived & VendRcvdDetail
                        string insertRcvSql = @"
                            INSERT INTO VendReceived (VendID, DT, RecieptID, UserID, ProcessID, Issuance_RefID, UserName, MachineName)
                            VALUES (0, GETDATE(), '', 0, @ProcessID, 0, @UserName, @MachineName);
                            SELECT CAST(SCOPE_IDENTITY() as int);";
                        int rcvRefId = await db.ExecuteScalarAsync<int>(insertRcvSql, request, trans);

                        // Get NextProcessID
                        string nextProcSql = @"
                            SELECT TOP 1 ProcessID 
                            FROM VItemProcesses 
                            WHERE IsExist = @ItemID 
                              AND ItemSNo > ISNULL((SELECT TOP 1 ItemSNo FROM VItemProcesses WHERE IsExist = @ItemID AND ProcessID = @ProcessID), 0)
                            ORDER BY ItemSNo";
                        int? nextProcId = await db.ExecuteScalarAsync<int?>(nextProcSql, new { ItemID = request.ItemID, ProcessID = request.ProcessID }, trans);

                        string insertRcvDetailSql = @"
                            INSERT INTO VendRcvdDetail 
                                (RefID, ItemCode, RecieptID, RcvdQty, Rate, LotNo, NextProcessID, OrderNo, CountedBy, Issue_RefID, ProcessID, RcvdWeight, Opening_RefID)
                            VALUES 
                                (@RefID, @ItemID, '', @Qty, 0, '0', @NextProcessID, 'Stock-Order', '', 0, @ProcessID, 0, @Opening_RefID)";

                        await db.ExecuteAsync(insertRcvDetailSql, new
                        {
                            RefID = rcvRefId,
                            ItemID = request.ItemID,
                            Qty = request.Qty,
                            NextProcessID = nextProcId ?? 0,
                            ProcessID = request.ProcessID,
                            Opening_RefID = sooEntryId
                        }, trans);

                        trans.Commit();
                        return sooEntryId;
                    }
                    catch
                    {
                        trans.Rollback();
                        throw;
                    }
                }
            }
        }

        public async Task UpdateRemarksAsync(int entryId, string remarks)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "UPDATE StockOrderOpening SET Remarks = @Remarks WHERE EntryID = @EntryID";
                await db.ExecuteAsync(sql, new { EntryID = entryId, Remarks = remarks });
            }
        }

        public async Task<bool> DeleteOpeningAsync(int entryId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();

                // Verify if issuance has been made
                string checkSql = @"
                    SELECT ISNULL(SUM(IssQty), 0) 
                    FROM VendRcvdDetail 
                    WHERE Opening_RefID = @EntryID";
                int issQty = await db.ExecuteScalarAsync<int>(checkSql, new { EntryID = entryId });

                if (issQty > 0)
                {
                    return false; // Can't delete, issuance exists
                }

                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        string delDetailSql = "DELETE FROM VendRcvdDetail WHERE Opening_RefID = @EntryID";
                        await db.ExecuteAsync(delDetailSql, new { EntryID = entryId }, trans);

                        string delRcvSql = "DELETE FROM VendReceived WHERE EntryID NOT IN (SELECT RefID FROM VendRcvdDetail WHERE RefID IS NOT NULL)";
                        await db.ExecuteAsync(delRcvSql, null, trans);

                        string delSOOSql = "DELETE FROM StockOrderOpening WHERE EntryID = @EntryID";
                        await db.ExecuteAsync(delSOOSql, new { EntryID = entryId }, trans);

                        trans.Commit();
                        return true;
                    }
                    catch
                    {
                        trans.Rollback();
                        throw;
                    }
                }
            }
        }
    }
}
