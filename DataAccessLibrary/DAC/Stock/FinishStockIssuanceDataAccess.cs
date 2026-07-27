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
    public class FinishStockIssuanceDataAccess : IFinishStockIssuanceDataAccess
    {
        private readonly string _connectionString;

        public FinishStockIssuanceDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<string> GetNextIssuanceNoAsync(DateTime date)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string dateStr = date.ToString("MM/dd/yy");
                string prefix = "ISS-" + date.ToString("yyMMdd-");

                string sql = @"
                    SELECT MAX(CAST(RIGHT(IssNo, 3) AS INT)) 
                    FROM IssItemsSimple 
                    WHERE CONVERT(CHAR(8), DT, 1) = @DateStr";

                var maxVal = await db.ExecuteScalarAsync<int?>(sql, new { DateStr = dateStr });
                int nextSeq = (maxVal ?? 0) + 1;
                return prefix + nextSeq.ToString("D3");
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

        public async Task<List<SFTargetStoreModel>> GetAvailableStoresAsync(string itemId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DISTINCT Store_RefID AS EntryID, StoreName
                    FROM VItemsWithLotWiseStock
                    WHERE ItemID = @ItemId AND NetQty > 0
                    ORDER BY StoreName";
                var results = await db.QueryAsync<SFTargetStoreModel>(sql, new { ItemId = itemId });
                return results.AsList();
            }
        }

        public async Task<List<LocationLotStockModel>> GetAvailableLocationsAsync(string itemId, int storeRefId, string custCode)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, ItemID, ISNULL(LotNo, '') AS LotNo, ISNULL(Batch_No, '') AS Batch_No,
                           NetQty, Shelf_RefID, Store_RefID, ISNULL(StoreName, '') AS StoreName,
                           ISNULL(RackNo, '') AS RackNo, ISNULL(ShelfNo, '') AS ShelfNo, ISNULL(Remarks, '') AS Remarks
                    FROM VItemsWithLotWiseStock
                    WHERE ItemID = @ItemId AND Store_RefID = @StoreRefId AND NetQty > 0
                    ORDER BY EntryID";
                var results = await db.QueryAsync<LocationLotStockModel>(sql, new { ItemId = itemId, StoreRefId = storeRefId });
                return results.AsList();
            }
        }

        public async Task<bool> SaveIssuanceAsync(FinishStockIssuanceMasterModel master, List<FinishStockIssuanceItemModel> items)
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
                        // Generate sequential IssNo inside transaction
                        string dateStr = master.DT.ToString("MM/dd/yy");
                        string prefix = "ISS-" + master.DT.ToString("yyMMdd-");
                        string seqSql = @"
                            SELECT MAX(CAST(RIGHT(IssNo, 3) AS INT)) 
                            FROM IssItemsSimple WITH (UPDLOCK, HOLDLOCK)
                            WHERE CONVERT(CHAR(8), DT, 1) = @DateStr";
                        var maxVal = await db.ExecuteScalarAsync<int?>(seqSql, new { DateStr = dateStr }, trans);
                        int nextSeq = (maxVal ?? 0) + 1;
                        master.IssNo = prefix + nextSeq.ToString("D3");

                        string masterSql = @"
                            INSERT INTO IssItemsSimple(IssNo, DT, OrderNo, Department, AttnPerson, EntryDT, UserName, ComputerName, SampleIssuance, CustCode)
                            VALUES(@IssNo, @DT, @OrderNo, @Department, @AttnPerson, GETDATE(), @UserName, 'WEB', 0, @CustCode)";

                        await db.ExecuteAsync(masterSql, master, trans);

                        foreach (var item in items)
                        {
                            // 1. Insert into ItemLedgerSimple
                            string desc = $"Item Issued To {master.AttnPerson} Of {master.Department}";
                            string ledgerSql = @"
                                INSERT INTO ItemLedgerSimple(ItemID, DT, Description, Qty, Balance, OrderNo, Location, Remarks)
                                VALUES(@ItemID, @DT, @Description, @Qty, 0, @OrderNo, @Location, @Remarks);
                                SELECT CAST(SCOPE_IDENTITY() as int);";

                            int ledgerRefId = await db.ExecuteScalarAsync<int>(ledgerSql, new
                            {
                                ItemID = item.ItemID,
                                DT = master.DT,
                                Description = desc,
                                Qty = -item.Qty,
                                OrderNo = master.OrderNo,
                                Location = item.LocationDisplay,
                                Remarks = item.Remarks
                            }, trans);

                            // 2. Insert into IssItemsSimpleDetail
                            string detailSql = @"
                                INSERT INTO IssItemsSimpleDetail(IssNo, ItemID, Qty, InHand, LedgerRefID, OrderNo, Location, Remarks, LotNo_Manual, BatchNo_Manual, RCV_ISD_RefID)
                                VALUES(@IssNo, @ItemID, @Qty, @InHand, @LedgerRefID, @OrderNo, @Location, @Remarks, @LotNo_Manual, @BatchNo_Manual, @RCV_ISD_RefID);
                                SELECT CAST(SCOPE_IDENTITY() as int);";

                            int detailId = await db.ExecuteScalarAsync<int>(detailSql, new
                            {
                                IssNo = master.IssNo,
                                ItemID = item.ItemID,
                                Qty = item.Qty,
                                InHand = item.AvailableQty - item.Qty,
                                LedgerRefID = ledgerRefId,
                                OrderNo = master.OrderNo,
                                Location = item.LocationDisplay,
                                Remarks = item.Remarks,
                                LotNo_Manual = item.LotNo,
                                BatchNo_Manual = item.BatchNo,
                                RCV_ISD_RefID = item.RISD_EntryID
                            }, trans);

                            // 3. Insert into IssItemsSimpleDetail_More
                            string moreSql = @"
                                INSERT INTO IssItemsSimpleDetail_More(IISD_RefID, IssdQty, Shelf_RefID, Remarks)
                                VALUES(@IISD_RefID, @IssdQty, @Shelf_RefID, @Remarks)";

                            await db.ExecuteAsync(moreSql, new
                            {
                                IISD_RefID = detailId,
                                IssdQty = (int)item.Qty,
                                Shelf_RefID = item.Shelf_RefID,
                                Remarks = item.Remarks
                            }, trans);

                            // 4. Update location stock via SP
                            await db.ExecuteAsync(
                                "EXEC UpdateItemLocation_SP @MaterialID = @ItemID, @LocationID = @ShelfID, @Qty = @Qty, @Sample_Type = -1, @CustCode = @CustCode",
                                new
                                {
                                    ItemID = item.ItemID,
                                    ShelfID = item.Shelf_RefID,
                                    Qty = item.Qty,
                                    CustCode = master.CustCode
                                },
                                trans,
                                commandType: CommandType.Text);
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
