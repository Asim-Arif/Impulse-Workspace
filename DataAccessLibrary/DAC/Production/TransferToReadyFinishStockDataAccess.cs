using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Production
{
    public class TransferToReadyFinishStockDataAccess : ITransferToReadyFinishStockDataAccess
    {
        private readonly IConfiguration _config;

        public TransferToReadyFinishStockDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("ImpulseConnection")
            ?? string.Empty;

        public async Task<List<LookupItemInt>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT 0 AS Id, '<All Makers>' AS Name
                UNION ALL
                SELECT VendID AS Id, ISNULL('{ ' + VendID1 + ' } ' + VenderName, VenderName) AS Name
                FROM Makers
                ORDER BY Name";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetArticlesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT ItemID AS Id,
                       ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }' AS Name
                FROM Items
                WHERE InActive = 0
                ORDER BY ItemName";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<StoreLookupDto>> GetStoresAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT EntryID, StoreName FROM Stores ORDER BY StoreName";
            return (await db.QueryAsync<StoreLookupDto>(sql)).ToList();
        }

        public async Task<List<ShelfLookupDto>> GetShelvesByStoreAsync(int storeId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT EntryID, Store_RefID, RackNo, ShelfNo
                FROM VStoreShelfs
                WHERE Store_RefID = @Store_RefID
                ORDER BY RackNo, ShelfNo";
            return (await db.QueryAsync<ShelfLookupDto>(sql, new { Store_RefID = storeId })).ToList();
        }

        public async Task<string> GetShelfRemarksAsync(string itemId, int shelfRefId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT TOP 1 Remarks
                FROM RcvItemsSimpleDetail_Placement
                WHERE Shelf_RefID = @Shelf_RefID
                  AND RISD_RefID IN (SELECT EntryID FROM RcvItemsSimpleDetail WHERE ItemID = @ItemID)
                ORDER BY EntryID DESC";

            return await db.ExecuteScalarAsync<string>(sql, new { ItemID = itemId, Shelf_RefID = shelfRefId }) ?? string.Empty;
        }

        public async Task<List<ReadyFinishStockAvailableItem>> GetAvailableLotsAsync(ReadyFinishStockFilter filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            string sql = @"
                SELECT VRD_EntryID, VendID, VendID1, VenderName, RecieptID, DT,
                       Description, ItemCode, ItemName, ItemSize, SizeUnit,
                       RcvdQty, ISNULL(Wastage, 0) AS Wastage, ISNULL(LostQty, 0) AS LostQty,
                       ISNULL(IssQty, 0) AS IssQty, ISNULL(ReWorkQty, 0) AS ReWorkQty,
                       OrderNo, LotNo, Batch_No, Mill_Certificate_No
                FROM VVendReceivingList
                WHERE OnlyDT >= @DtFrom AND OnlyDT <= @DtTo
                  AND NextProcessID IS NULL
                  AND (RcvdQty - ISNULL(Wastage,0) - ISNULL(LostQty,0)) > IssQty
                  AND ReqAuth = 0";

            var p = new DynamicParameters();
            p.Add("@DtFrom", filter.DtFrom.Date);
            p.Add("@DtTo", filter.DtTo.Date.AddDays(1).AddTicks(-1));

            if (filter.MakerId > 0)
            {
                sql += " AND VendID = @VendID";
                p.Add("@VendID", filter.MakerId);
            }

            if (!string.IsNullOrWhiteSpace(filter.ArticleId) && filter.ArticleId != "0")
            {
                sql += " AND ItemCode = @ItemCode";
                p.Add("@ItemCode", filter.ArticleId.Trim());
            }

            if (!string.IsNullOrWhiteSpace(filter.LotNo))
            {
                sql += " AND LotNo LIKE @LotNo";
                p.Add("@LotNo", $"%{filter.LotNo.Trim()}%");
            }

            sql += " ORDER BY DT";

            var list = (await db.QueryAsync<ReadyFinishStockAvailableItem>(sql, p)).ToList();
            foreach (var item in list)
            {
                item.TransferQty = item.AvailQty;
            }

            return list;
        }

        private async Task<string> GenerateRcvNoAsync(IDbConnection db, IDbTransaction tx)
        {
            string prefix = $"RCV-{DateTime.Today:yyMMdd}-";
            const string sql = @"
                SELECT MAX(CAST(RIGHT(RcvNo, LEN(RcvNo) - 13) AS INT))
                FROM RcvItemsSimple
                WHERE RcvNo LIKE @PrefixPattern";

            int maxSeq = await db.ExecuteScalarAsync<int?>(sql, new { PrefixPattern = $"{prefix}%" }, tx) ?? 0;
            return $"{prefix}{(maxSeq + 1):D3}";
        }

        private async Task<decimal> GetItemBalanceAsync(IDbConnection db, IDbTransaction tx, string itemId, DateTime date)
        {
            const string sql = @"
                SELECT ISNULL(SUM(Qty), 0)
                FROM ItemLedgerSimple
                WHERE ItemID = @ItemID AND DT <= @DT";

            return await db.ExecuteScalarAsync<decimal>(sql, new { ItemID = itemId, DT = date }, tx);
        }

        public async Task<PostTransferResult> PostTransferToReadyFinishAsync(PostTransferRequest request)
        {
            if (request.Items == null || !request.Items.Any())
            {
                return new PostTransferResult { Success = false, Message = "No lots selected for transfer." };
            }

            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                // 1. Insert TransferredToReadyFinishLots Header
                await db.ExecuteAsync(@"
                    INSERT INTO TransferredToReadyFinishLots (UserName, MachineName, DT, ItemID, Remarks)
                    VALUES (@UserName, @MachineName, @DT, @ItemID, @Remarks)",
                    new
                    {
                        UserName = request.UserName,
                        MachineName = request.MachineName,
                        DT = DateTime.Today,
                        ItemID = request.ArticleId,
                        Remarks = request.Remarks
                    }, tx);

                long lEntryID = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM TransferredToReadyFinishLots", transaction: tx);

                // 2. Generate RcvNo
                string strRcvNo = await GenerateRcvNoAsync(db, tx);

                string strPOs = string.Join(",", request.Items.Select(x => x.OrderNo).Distinct());
                string strLots = string.Join(",", request.Items.Select(x => x.LotNo).Distinct());
                string strRcvdFrom = $"From Production Lot No.: {strLots}";

                // 3. Insert RcvItemsSimple
                await db.ExecuteAsync(@"
                    INSERT INTO RcvItemsSimple (RcvNo, DT, OrderNo, RcvdBy, RcvdFrom, UserName, ComputerName)
                    VALUES (@RcvNo, @DT, @OrderNo, @RcvdBy, @RcvdFrom, @UserName, @ComputerName)",
                    new
                    {
                        RcvNo = strRcvNo,
                        DT = DateTime.Now,
                        OrderNo = strPOs,
                        RcvdBy = request.UserName,
                        RcvdFrom = strRcvdFrom,
                        UserName = request.UserName,
                        ComputerName = request.MachineName
                    }, tx);

                // 4. Process each transferred lot
                foreach (var item in request.Items)
                {
                    decimal dQty = item.TransferQty;

                    // Detail in TransferredToReadyFinishLotsDetail
                    await db.ExecuteAsync(@"
                        INSERT INTO TransferredToReadyFinishLotsDetail (RefID, VRD_RefID, Qty)
                        VALUES (@RefID, @VRD_RefID, @Qty)",
                        new
                        {
                            RefID = lEntryID,
                            VRD_RefID = item.VRD_EntryID,
                            Qty = (int)dQty
                        }, tx);

                    long lTTRFD_EntryID = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM TransferredToReadyFinishLotsDetail", transaction: tx);

                    // Update VendRcvdDetail.IssQty
                    await db.ExecuteAsync(@"
                        UPDATE VendRcvdDetail
                        SET IssQty = IssQty + @Qty
                        WHERE EntryID = @VRD_EntryID",
                        new { Qty = dQty, VRD_EntryID = item.VRD_EntryID }, tx);

                    // Balance for ItemLedgerSimple
                    decimal dBal = await GetItemBalanceAsync(db, tx, request.ArticleId, DateTime.Today);
                    dBal += dQty;

                    // Insert ItemLedgerSimple
                    await db.ExecuteAsync(@"
                        INSERT INTO ItemLedgerSimple (ItemID, DT, Description, Qty, Balance, OrderNo, Location, Remarks)
                        VALUES (@ItemID, @DT, @Description, @Qty, @Balance, @OrderNo, @Location, @Remarks)",
                        new
                        {
                            ItemID = request.ArticleId,
                            DT = DateTime.Today,
                            Description = request.Remarks,
                            Qty = (double)dQty,
                            Balance = (double)dBal,
                            OrderNo = item.OrderNo,
                            Location = request.ShelfLocationText,
                            Remarks = item.LotNo
                        }, tx);

                    long ledgerRefId = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM ItemLedgerSimple", transaction: tx);

                    // Insert RcvItemsSimpleDetail
                    await db.ExecuteAsync(@"
                        INSERT INTO RcvItemsSimpleDetail (
                            RcvNo, ItemID, Qty, InHand, LedgerRefID, OrderNo,
                            Location, Remarks, TTRFLD_RefID, LotNo_Manual, BatchNo_Manual, Mill_Certificate_No
                        )
                        VALUES (
                            @RcvNo, @ItemID, @Qty, @InHand, @LedgerRefID, @OrderNo,
                            @Location, @Remarks, @TTRFLD_RefID, @LotNo_Manual, @BatchNo_Manual, @Mill_Certificate_No
                        )",
                        new
                        {
                            RcvNo = strRcvNo,
                            ItemID = request.ArticleId,
                            Qty = (double)dQty,
                            InHand = (double)dBal,
                            LedgerRefID = ledgerRefId,
                            OrderNo = item.OrderNo,
                            Location = request.ShelfLocationText,
                            Remarks = request.Remarks,
                            TTRFLD_RefID = lTTRFD_EntryID,
                            LotNo_Manual = item.LotNo,
                            BatchNo_Manual = item.BatchNo,
                            Mill_Certificate_No = item.MillCertNo
                        }, tx);

                    long lRISD_EntryID = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM RcvItemsSimpleDetail", transaction: tx);

                    // Insert RcvItemsSimpleDetail_Placement
                    await db.ExecuteAsync(@"
                        INSERT INTO RcvItemsSimpleDetail_Placement (RISD_RefID, Shelf_RefID, RcvdQty, Remarks)
                        VALUES (@RISD_RefID, @Shelf_RefID, @RcvdQty, @Remarks)",
                        new
                        {
                            RISD_RefID = lRISD_EntryID,
                            Shelf_RefID = request.SelectedShelfId,
                            RcvdQty = (int)dQty,
                            Remarks = request.Remarks
                        }, tx);
                }

                tx.Commit();

                return new PostTransferResult
                {
                    Success = true,
                    RcvNo = strRcvNo,
                    TransferEntryID = lEntryID,
                    Message = $"Transferred successfully under Receipt #{strRcvNo}."
                };
            }
            catch (Exception ex)
            {
                tx.Rollback();
                return new PostTransferResult
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }
    }
}
