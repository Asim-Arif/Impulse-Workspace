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
    public class FinishMovementDataAccess : IFinishMovementDataAccess
    {
        private readonly string _connectionString;

        public FinishMovementDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
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

        public async Task<List<FinishFromLocationModel>> GetFromLocationsAsync(string itemId)
        {
            if (string.IsNullOrWhiteSpace(itemId)) return new List<FinishFromLocationModel>();

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ISNULL(EntryID, 0) AS EntryID, ISNULL(Shelf_RefID, 0) AS Shelf_RefID, 
                           ISNULL(LotNo, '') AS LotNo, ISNULL(Batch_No, '') AS Batch_No, 
                           ISNULL(Mill_Certificate_No, '') AS Mill_Certificate_No,
                           ISNULL(StoreName, '') AS StoreName, ISNULL(RackNo, '') AS RackNo, 
                           ISNULL(ShelfNo, '') AS ShelfNo, ISNULL(NetQty, 0) AS NetQty,
                           ISNULL(Remarks, '') AS Remarks
                    FROM VItemsWithLotWiseStock
                    WHERE ItemID = @ItemID AND NetQty > 0
                    ORDER BY LotNo, Batch_No";
                var results = await db.QueryAsync<FinishFromLocationModel>(sql, new { ItemID = itemId });
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

        public async Task<List<SFTargetRackModel>> GetRacksAsync(int storeRefId)
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

        public async Task<List<SFTargetShelfModel>> GetShelvesAsync(int storeRefId, string rackNo)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, Store_RefID, ISNULL(RackNo, '') AS RackNo, ISNULL(ShelfNo, '') AS ShelfNo, ISNULL(StoreName, '') AS StoreName
                    FROM VStoreShelfs
                    WHERE Store_RefID = @StoreRefId AND ISNULL(RackNo, '') = @RackNo
                    ORDER BY ShelfNo";
                var results = await db.QueryAsync<SFTargetShelfModel>(sql, new { StoreRefId = storeRefId, RackNo = rackNo });
                return results.AsList();
            }
        }

        public async Task<double> GetTargetShelfQtyAsync(string itemId, int shelfRefId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ISNULL(SUM(NetQty), 0)
                    FROM VItemsWithShelfWiseStock
                    WHERE Shelf_RefID = @ShelfRefId AND ItemID = @ItemID";
                return await db.ExecuteScalarAsync<double>(sql, new { ShelfRefId = shelfRefId, ItemID = itemId });
            }
        }

        public async Task<bool> ExecuteMovementAsync(FinishMovementRequest request)
        {
            if (request == null || request.SourceLocation == null || request.MoveQty <= 0)
                return false;

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        DateTime today = DateTime.Today;
                        string dateStr = today.ToString("MM/dd/yy");

                        // 1. Generate sequential IssNo for deduction
                        string issPrefix = "ISS-" + today.ToString("yyMMdd-");
                        string issSeqSql = @"
                            SELECT MAX(CAST(RIGHT(IssNo, 3) AS INT)) 
                            FROM IssItemsSimple WITH (UPDLOCK, HOLDLOCK)
                            WHERE CONVERT(CHAR(8), DT, 1) = @DateStr";
                        var maxIss = await db.ExecuteScalarAsync<int?>(issSeqSql, new { DateStr = dateStr }, trans);
                        string issNo = issPrefix + ((maxIss ?? 0) + 1).ToString("D3");

                        // 2. Generate sequential RcvNo for addition
                        string rcvPrefix = "RCV-" + today.ToString("yyMMdd-");
                        string rcvSeqSql = @"
                            SELECT MAX(CAST(RIGHT(RcvNo, 3) AS INT)) 
                            FROM RcvItemsSimple WITH (UPDLOCK, HOLDLOCK)
                            WHERE CONVERT(CHAR(8), DT, 1) = @DateStr";
                        var maxRcv = await db.ExecuteScalarAsync<int?>(rcvSeqSql, new { DateStr = dateStr }, trans);
                        string rcvNo = rcvPrefix + ((maxRcv ?? 0) + 1).ToString("D3");

                        string sourceLocDisplay = $"{request.SourceLocation.StoreName}[{request.SourceLocation.RackNo}][{request.SourceLocation.ShelfNo}]";

                        // --- A. Material Issuance (Deduct from Source Shelf) ---
                        string issMasterSql = @"
                            INSERT INTO IssItemsSimple(IssNo, DT, OrderNo, Department, AttnPerson, EntryDT, UserName, ComputerName, SampleIssuance)
                            VALUES(@IssNo, GETDATE(), 'For Location Change', '', '', GETDATE(), @UserName, 'WEB', 0)";
                        await db.ExecuteAsync(issMasterSql, new { IssNo = issNo, UserName = request.UserName }, trans);

                        string issLedgerSql = @"
                            INSERT INTO ItemLedgerSimple(ItemID, DT, Description, Qty, Balance, OrderNo, Location, Remarks)
                            VALUES(@ItemID, GETDATE(), 'For Material Location Change', @Qty, 0, '', @Location, '');
                            SELECT CAST(SCOPE_IDENTITY() as int);";
                        int issLedgerId = await db.ExecuteScalarAsync<int>(issLedgerSql, new
                        {
                            ItemID = request.ItemID,
                            Qty = -request.MoveQty,
                            Location = sourceLocDisplay
                        }, trans);

                        string issDetailSql = @"
                            INSERT INTO IssItemsSimpleDetail(IssNo, ItemID, Qty, InHand, LedgerRefID, OrderNo, Location, Remarks, LotNo_Manual, BatchNo_Manual, RCV_ISD_RefID)
                            VALUES(@IssNo, @ItemID, @Qty, 0, @LedgerRefID, '', @Location, '', @LotNo, @BatchNo, @SourceEntryID);
                            SELECT CAST(SCOPE_IDENTITY() as int);";
                        int iisdEntryId = await db.ExecuteScalarAsync<int>(issDetailSql, new
                        {
                            IssNo = issNo,
                            ItemID = request.ItemID,
                            Qty = request.MoveQty,
                            LedgerRefID = issLedgerId,
                            Location = sourceLocDisplay,
                            LotNo = request.SourceLocation.LotNo,
                            BatchNo = request.SourceLocation.Batch_No,
                            SourceEntryID = request.SourceLocation.EntryID
                        }, trans);

                        string issMoreSql = @"
                            INSERT INTO IssItemsSimpleDetail_More(IISD_RefID, IssdQty, Shelf_RefID, Remarks)
                            VALUES(@IISD_RefID, @IssdQty, @Shelf_RefID, '')";
                        await db.ExecuteAsync(issMoreSql, new
                        {
                            IISD_RefID = iisdEntryId,
                            IssdQty = request.MoveQty,
                            Shelf_RefID = request.SourceLocation.Shelf_RefID
                        }, trans);

                        // --- B. Material Receiving (Add to Target Shelf) ---
                        string rcvMasterSql = @"
                            INSERT INTO RcvItemsSimple(RcvNo, DT, OrderNo, RcvdBy, RcvdFrom, EntryDT, UserName, ComputerName)
                            VALUES(@RcvNo, GETDATE(), 'Material Movement', '', '', GETDATE(), @UserName, 'WEB')";
                        await db.ExecuteAsync(rcvMasterSql, new { RcvNo = rcvNo, UserName = request.UserName }, trans);

                        string rcvLedgerSql = @"
                            INSERT INTO ItemLedgerSimple(ItemID, DT, Description, Qty, Balance, OrderNo, Location, Remarks)
                            VALUES(@ItemID, GETDATE(), 'Material Movement', @Qty, 0, '', @Location, '');
                            SELECT CAST(SCOPE_IDENTITY() as int);";
                        int rcvLedgerId = await db.ExecuteScalarAsync<int>(rcvLedgerSql, new
                        {
                            ItemID = request.ItemID,
                            Qty = request.MoveQty,
                            Location = request.TargetLocationDisplay
                        }, trans);

                        string rcvDetailSql = @"
                            INSERT INTO RcvItemsSimpleDetail(RcvNo, ItemID, Qty, InHand, LedgerRefID, OrderNo, Location, Remarks, LotNo_Manual, BatchNo_Manual, Mill_Certificate_No)
                            VALUES(@RcvNo, @ItemID, @Qty, 0, @LedgerRefID, '', '', '', @LotNo, @BatchNo, @MillCertNo);
                            SELECT CAST(SCOPE_IDENTITY() as int);";
                        int risdEntryId = await db.ExecuteScalarAsync<int>(rcvDetailSql, new
                        {
                            RcvNo = rcvNo,
                            ItemID = request.ItemID,
                            Qty = request.MoveQty,
                            LedgerRefID = rcvLedgerId,
                            LotNo = request.SourceLocation.LotNo,
                            BatchNo = request.SourceLocation.Batch_No,
                            MillCertNo = request.SourceLocation.Mill_Certificate_No
                        }, trans);

                        string rcvPlacementSql = @"
                            INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID, Shelf_RefID, RcvdQty, Remarks)
                            VALUES(@RISD_RefID, @Shelf_RefID, @RcvdQty, @Remarks)";
                        await db.ExecuteAsync(rcvPlacementSql, new
                        {
                            RISD_RefID = risdEntryId,
                            Shelf_RefID = request.TargetShelfRefId,
                            RcvdQty = request.MoveQty,
                            Remarks = request.SourceLocation.Remarks
                        }, trans);

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
