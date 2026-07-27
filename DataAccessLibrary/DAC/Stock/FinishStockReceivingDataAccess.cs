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
    public class FinishStockReceivingDataAccess : IFinishStockReceivingDataAccess
    {
        private readonly string _connectionString;

        public FinishStockReceivingDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<string> GetNextReceivingNoAsync(DateTime date)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string dateStr = date.ToString("MM/dd/yy");
                string prefix = "RCV-" + date.ToString("yyMMdd-");

                string sql = @"
                    SELECT MAX(CAST(RIGHT(RcvNo, 3) AS INT)) 
                    FROM RcvItemsSimple 
                    WHERE CONVERT(CHAR(8), DT, 1) = @DateStr";

                var maxVal = await db.ExecuteScalarAsync<int?>(sql, new { DateStr = dateStr });
                int nextSeq = (maxVal ?? 0) + 1;
                return prefix + nextSeq.ToString("D3");
            }
        }

        public async Task<string> GetNextLotNoAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string yearPrefix = DateTime.Today.ToString("yy");
                string sql = @"
                    SELECT MAX(CAST(LotNo AS BIGINT))
                    FROM Lots_List
                    WHERE LotNo LIKE @Prefix + '%' AND ISNUMERIC(LotNo) = 1";

                var maxLot = await db.ExecuteScalarAsync<long?>(sql, new { Prefix = yearPrefix });
                if (maxLot.HasValue)
                {
                    return (maxLot.Value + 1).ToString();
                }
                return yearPrefix + "00001";
            }
        }

        public async Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT CustCode, ISNULL(Country, '') AS Country 
                    FROM ForeignCustomers 
                    ORDER BY CustCode";
                var list = (await db.QueryAsync<CustomerLookupModel>(sql)).AsList();
                list.Insert(0, new CustomerLookupModel { CustCode = "Stock", Country = "Local Company Stock" });
                return list;
            }
        }

        public async Task<List<ItemViewModel>> GetArticlesAsync(string custCode)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql;
                if (string.IsNullOrEmpty(custCode) || custCode == "Stock" || custCode == "0")
                {
                    sql = @"
                        SELECT ItemID, '{' + ItemID + '} ' + ItemName AS ItemName, ISNULL(Unit, '') AS Unit
                        FROM Items
                        WHERE ISNULL(InActive, 0) = 0
                        ORDER BY ItemName";
                }
                else
                {
                    sql = @"
                        SELECT CompItemID AS ItemID, '{' + CompItemID + '} ' + ItemID + ' ' + ISNULL(CompDesc, '') AS ItemName, ISNULL(Unit, '') AS Unit
                        FROM VFCustomerCatalog
                        WHERE CustCode = @CustCode
                        ORDER BY CompItemID";
                }
                var results = await db.QueryAsync<ItemViewModel>(sql, new { CustCode = custCode });
                return results.AsList();
            }
        }

        public async Task<List<SFTargetStoreModel>> GetStoresAsync(string userName)
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

        public async Task<List<SFTargetShelfModel>> GetShelvesAsync(int storeRefId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, Store_RefID, ISNULL(RackNo, '') AS RackNo, ISNULL(ShelfNo, '') AS ShelfNo, ISNULL(StoreName, '') AS StoreName
                    FROM VStoreShelfs
                    WHERE Store_RefID = @StoreRefId
                    ORDER BY RackNo, ShelfNo";
                var results = await db.QueryAsync<SFTargetShelfModel>(sql, new { StoreRefId = storeRefId });
                return results.AsList();
            }
        }

        public async Task<bool> IsLotNoExistsAsync(string lotNo)
        {
            if (string.IsNullOrWhiteSpace(lotNo)) return false;
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT COUNT(1) FROM RcvItemsSimpleDetail WHERE LotNo_Manual = @LotNo";
                int count = await db.ExecuteScalarAsync<int>(sql, new { LotNo = lotNo });
                return count > 0;
            }
        }

        public async Task<bool> IsBatchNoExistsAsync(string batchNo)
        {
            if (string.IsNullOrWhiteSpace(batchNo)) return false;
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT COUNT(1) FROM RcvItemsSimpleDetail WHERE BatchNo_Manual = @BatchNo";
                int count = await db.ExecuteScalarAsync<int>(sql, new { BatchNo = batchNo });
                return count > 0;
            }
        }

        public async Task<bool> SaveReceivingAsync(FinishStockReceivingMasterModel master, List<FinishStockReceivingItemModel> items)
        {
            if (master == null || items == null || !items.Any())
                return false;

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        // Generate sequential RcvNo inside transaction
                        string dateStr = master.DT.ToString("MM/dd/yy");
                        string prefix = "RCV-" + master.DT.ToString("yyMMdd-");
                        string seqSql = @"
                            SELECT MAX(CAST(RIGHT(RcvNo, 3) AS INT)) 
                            FROM RcvItemsSimple WITH (UPDLOCK, HOLDLOCK)
                            WHERE CONVERT(CHAR(8), DT, 1) = @DateStr";
                        var maxVal = await db.ExecuteScalarAsync<int?>(seqSql, new { DateStr = dateStr }, trans);
                        int nextSeq = (maxVal ?? 0) + 1;
                        master.RcvNo = prefix + nextSeq.ToString("D3");

                        string masterSql = @"
                            INSERT INTO RcvItemsSimple(RcvNo, DT, OrderNo, RcvdBy, RcvdFrom, EntryDT, UserName, ComputerName, CustCode)
                            VALUES(@RcvNo, @DT, @OrderNo, @RcvdBy, @RcvdFrom, GETDATE(), @UserName, 'WEB', @CustCode)";

                        await db.ExecuteAsync(masterSql, master, trans);

                        foreach (var item in items)
                        {
                            // 1. Insert into ItemLedgerSimple (positive Qty for Receiving)
                            string desc = $"Item Received From {master.RcvdFrom} Of {master.RcvdBy}";
                            string ledgerSql = @"
                                INSERT INTO ItemLedgerSimple(ItemID, DT, Description, Qty, Balance, OrderNo, Location, Remarks)
                                VALUES(@ItemID, @DT, @Description, @Qty, 0, @OrderNo, @Location, @Remarks);
                                SELECT CAST(SCOPE_IDENTITY() as int);";

                            int ledgerRefId = await db.ExecuteScalarAsync<int>(ledgerSql, new
                            {
                                ItemID = item.ItemID,
                                DT = master.DT,
                                Description = desc,
                                Qty = item.Qty,
                                OrderNo = master.OrderNo,
                                Location = item.LocationDisplay,
                                Remarks = item.Remarks
                            }, trans);

                            // 2. Insert into RcvItemsSimpleDetail
                            string detailSql = @"
                                INSERT INTO RcvItemsSimpleDetail(RcvNo, ItemID, Qty, InHand, LedgerRefID, OrderNo, Location, Remarks, LotNo_Manual, BatchNo_Manual, Mill_Certificate_No)
                                VALUES(@RcvNo, @ItemID, @Qty, 0, @LedgerRefID, @OrderNo, @Location, @Remarks, @LotNo_Manual, @BatchNo_Manual, @Mill_Certificate_No);
                                SELECT CAST(SCOPE_IDENTITY() as int);";

                            int detailId = await db.ExecuteScalarAsync<int>(detailSql, new
                            {
                                RcvNo = master.RcvNo,
                                ItemID = item.ItemID,
                                Qty = item.Qty,
                                LedgerRefID = ledgerRefId,
                                OrderNo = master.OrderNo,
                                Location = item.LocationDisplay,
                                Remarks = item.Remarks,
                                LotNo_Manual = item.LotNo_Manual,
                                BatchNo_Manual = item.BatchNo_Manual,
                                Mill_Certificate_No = item.Mill_Certificate_No
                            }, trans);

                            // 3. Insert into Lots_List
                            string lotSql = @"
                                IF NOT EXISTS (SELECT 1 FROM Lots_List WHERE LotNo = @LotNo)
                                BEGIN
                                    INSERT INTO Lots_List(LotNo, ItemID, Lot_Type, Reference_LotNo, Batch_No, Mill_Certificate_No)
                                    VALUES(@LotNo, @ItemID, 2, '', @BatchNo, @MillCertNo)
                                END";

                            await db.ExecuteAsync(lotSql, new
                            {
                                LotNo = item.LotNo_Manual,
                                ItemID = item.ItemID,
                                BatchNo = item.BatchNo_Manual,
                                MillCertNo = item.Mill_Certificate_No
                            }, trans);

                            // 4. Insert into RcvItemsSimpleDetail_Placement
                            string placementSql = @"
                                INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID, Shelf_RefID, RcvdQty, Remarks)
                                VALUES(@RISD_RefID, @Shelf_RefID, @RcvdQty, @Remarks)";

                            await db.ExecuteAsync(placementSql, new
                            {
                                RISD_RefID = detailId,
                                Shelf_RefID = item.Shelf_RefID,
                                RcvdQty = (int)item.Qty,
                                Remarks = item.Remarks
                            }, trans);
                        }

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
