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
    public class MakerRcvListDataAccess : IMakerRcvListDataAccess
    {
        private readonly IConfiguration _config;

        public MakerRcvListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        // ─────────────────────────────────────────────────────────────
        // MAIN LIST QUERY (mirrors RefreshLV logic from legacy VB6)
        // ─────────────────────────────────────────────────────────────
        public async Task<(List<MakerRcvListItem> Items, string ReportSql)> GetListAsync(MakerRcvListFilter filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            var parameters = new DynamicParameters();

            // ── STEP 1: Date range ──────────────────────────────────
            // Date range is bypassed if LotNo or OrderNo is entered
            DateTime now = DateTime.Today;
            DateTime dtFrom = filter.DtFrom;
            DateTime dtTo = filter.DtTo;

            switch (filter.DateRangeIndex)
            {
                case 0: dtFrom = now; dtTo = now; break;
                case 1: dtFrom = now.AddDays(-15); dtTo = now; break;
                case 2: dtFrom = now.AddDays(-30); dtTo = now; break;
                case 3: dtFrom = now.AddDays(-60); dtTo = now; break;
                case 4: dtFrom = now.AddDays(-90); dtTo = now; break;
                case 5: break; // custom — use filter values as-is
            }

            parameters.Add("@DtFrom", dtFrom);
            parameters.Add("@DtTo", dtTo);

            string whereClause;

            // If Lot No or Order No entered, ignore date range (legacy behaviour)
            if (!string.IsNullOrWhiteSpace(filter.LotNo) || !string.IsNullOrWhiteSpace(filter.OrderNo))
            {
                whereClause = " WHERE OnlyDT BETWEEN '1901-01-01' AND '2090-01-01'";
            }
            else
            {
                whereClause = " WHERE OnlyDT BETWEEN @DtFrom AND @DtTo";
            }

            // Always exclude stock openings
            whereClause += " AND ISNULL(Opening_RefID, 0) = 0";

            // ── STEP 2: Maker filter ─────────────────────────────────
            if (filter.MakerIds != null && filter.MakerIds.Any())
            {
                whereClause += $" AND VendID IN ({string.Join(",", filter.MakerIds)})";
            }

            // ── STEP 3: Category filter ──────────────────────────────
            if (filter.ItemCatIds != null && filter.ItemCatIds.Any())
            {
                var catList = string.Join(",", filter.ItemCatIds.Select(c => $"'{c.Replace("'", "''")}'"));
                whereClause += $" AND CatID IN ({catList})";
            }

            // ── STEP 4: Group filter ─────────────────────────────────
            if (filter.ItemGroupIds != null && filter.ItemGroupIds.Any())
            {
                whereClause += $" AND GroupID IN ({string.Join(",", filter.ItemGroupIds)})";
            }

            // ── STEP 5: Article / Item filter ───────────────────────
            if (!string.IsNullOrWhiteSpace(filter.ItemId) && filter.ItemId != "0")
            {
                parameters.Add("@ItemId", filter.ItemId);
                whereClause += " AND ItemCode = @ItemId";
            }

            // ── STEP 6: Process filter ───────────────────────────────
            if (filter.ProcessIds != null && filter.ProcessIds.Any())
            {
                whereClause += $" AND ProcessID IN ({string.Join(",", filter.ProcessIds)})";
            }

            // ── STEP 7: Lot No. filter ───────────────────────────────
            if (!string.IsNullOrWhiteSpace(filter.LotNo))
            {
                parameters.Add("@LotNo", filter.LotNo.Trim());
                whereClause += " AND LotNo = @LotNo";
            }

            // ── STEP 8: Order No. filter ─────────────────────────────
            if (!string.IsNullOrWhiteSpace(filter.OrderNo))
            {
                parameters.Add("@OrderNo", filter.OrderNo.Trim());
                whereClause += " AND OrderNo = @OrderNo";
            }

            // ── STEP 9: Master PO No. filter ────────────────────────
            if (!string.IsNullOrWhiteSpace(filter.MasterPONo))
            {
                parameters.Add("@MasterPONo", filter.MasterPONo.Trim());
                whereClause += " AND MasterPONo = @MasterPONo";
            }

            // ── STEP 10: Show Master PO Only ─────────────────────────
            if (filter.ShowMasterPOOnly)
            {
                whereClause += " AND MasterPONo IS NOT NULL";
            }

            // ── STEP 11: Repair / Rejection checkboxes ───────────────
            if (filter.RepairLots && filter.RejectionLots)
            {
                whereClause += " AND (ReWorkQty > 0 OR Wastage > 0)";
            }
            else if (filter.RepairLots)
            {
                whereClause += " AND (ReWorkQty > 0 OR ReWorkLot = 1)";
            }
            else if (filter.RejectionLots)
            {
                whereClause += " AND Wastage > 0";
            }

            // ── STEP 12: Regular Lots Only ───────────────────────────
            if (filter.RegularLotsOnly)
            {
                whereClause += " AND ReWorkLot = 0";
            }

            // ── STEP 13: In-Active Days ──────────────────────────────
            if (int.TryParse(filter.InActiveDays, out int inActiveDays) && inActiveDays > 0)
            {
                DateTime inActiveDate = now.AddDays(-inActiveDays);
                parameters.Add("@InActiveDate", inActiveDate);
                whereClause += " AND (LotNo <> '0' AND OnlyDT <= @InActiveDate AND IssQty = 0)";
            }

            // ── STEP 14: Customer filter ─────────────────────────────
            if (filter.CustomerCodes != null && filter.CustomerCodes.Any())
            {
                var custList = string.Join(",", filter.CustomerCodes.Select(c => $"'{c.Replace("'", "''")}'"));
                whereClause += $" AND CustCode IN ({custList})";
            }

            // ── STEP 15: Received but not Issued ─────────────────────
            if (filter.RcvdButNotIssued)
            {
                whereClause += @" AND (VVendReceivingList.VRD_EntryID NOT IN (SELECT Rcvd_RefID FROM VendIssdDetail)
                                  AND LotNo IN (SELECT LotNo FROM VRunningLots_Simple)
                                  AND LotNo NOT IN (SELECT LotNo FROM Lots_Closed))";
            }

            // ── STEP 16: Billed / Non-Billed ─────────────────────────
            if (filter.BilledOnly)
            {
                whereClause += " AND (MPB_D_EntryID IS NOT NULL)";
            }
            else if (filter.NonBilledOnly)
            {
                whereClause += " AND (MPB_D_EntryID IS NULL)";
            }

            // ── STEP 17: Build final SQL (ShowLastOnly changes structure) ──
            string sql;
            if (filter.ShowLastOnly)
            {
                // Subquery JOIN: only the last receiving entry per lot (where RcvdQty > IssQty)
                sql = $@"SELECT VVendReceivingList.*
                         FROM VVendReceivingList
                         INNER JOIN (
                             SELECT MAX(EntryID) AS LastEntryID
                             FROM VendRcvdDetail
                             WHERE RcvdQty > IssQty
                             GROUP BY LotNo
                         ) T1 ON VVendReceivingList.VRD_EntryID = T1.LastEntryID
                         {whereClause}
                         ORDER BY DT";
            }
            else
            {
                sql = $"SELECT VVendReceivingList.* FROM VVendReceivingList {whereClause} ORDER BY DT";
            }

            var items = (await db.QueryAsync<MakerRcvListItem>(sql, parameters)).ToList();
            return (items, sql);
        }

        // ─────────────────────────────────────────────────────────────
        // LOOKUP QUERIES
        // ─────────────────────────────────────────────────────────────
        public async Task<List<LookupItemInt>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT VendID AS Id, ISNULL('{ ' + VendID1 + ' } ' + VenderName, VenderName) AS Name FROM Makers ORDER BY VenderName";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT CustCode AS Id, CustCode AS Name FROM ForeignCustomers ORDER BY CustCode";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetItemCategoriesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT CatID AS Id, Description AS Name FROM ItemCatagories ORDER BY Description";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<LookupItemInt>> GetItemGroupsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ID AS Id, Description AS Name FROM ItemGroups ORDER BY Description";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetItemsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ItemID AS Id, ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }' AS Name FROM Items WHERE InActive = 0 ORDER BY ItemName";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<LookupItemInt>> GetProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ProcessID AS Id, Description AS Name FROM Processes ORDER BY Description";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        // ─────────────────────────────────────────────────────────────
        // DELETE RECEIVING
        // Mirrors mnuDeleteRcving_Click: log to deletions, delete detail, clean up orphan header
        // ─────────────────────────────────────────────────────────────
        public async Task<bool> DeleteReceivingAsync(long vrdEntryId, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                var p = new DynamicParameters();
                p.Add("@VRD_EntryID", vrdEntryId);
                p.Add("@UserName", userName);
                p.Add("@MachineName", machineName);

                // 1. Log to deletions archive
                await db.ExecuteAsync(@"
                    INSERT INTO VendRcvdDetail_Deletions
                        (LotNo, ProcessID, DTRcving, RcvdQty, WastageQty, ReWorkQty,
                         UserName, MachineName, VRD_RefID, VR_RefID, ItemID, MakerID)
                    SELECT
                        vrd.LotNo, vrd.ProcessID, vr.DT, vrd.RcvdQty, vrd.Wastage, vrd.ReWorkQty,
                        @UserName, @MachineName, vrd.EntryID, vr.EntryID, vrd.ItemCode, vr.VendID
                    FROM VendReceived vr
                    INNER JOIN VendRcvdDetail vrd ON vr.EntryID = vrd.RefID
                    WHERE vrd.EntryID = @VRD_EntryID", p, tx);

                // 2. Delete detail row
                await db.ExecuteAsync(
                    "DELETE FROM VendRcvdDetail WHERE EntryID = @VRD_EntryID", p, tx);

                // 3. Clean up orphan VendReceived headers
                await db.ExecuteAsync(
                    "DELETE FROM VendReceived WHERE EntryID NOT IN (SELECT RefID FROM VendRcvdDetail)",
                    transaction: tx);

                tx.Commit();
                return true;
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // CLOSE LOT
        // Mirrors mnuCloseLot_Click: INSERT INTO Lots_Closed
        // ─────────────────────────────────────────────────────────────
        public async Task<bool> CloseLotAsync(string lotNo, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            // Check if already closed
            int existing = await db.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM Lots_Closed WHERE LotNo = @LotNo",
                new { LotNo = lotNo });

            if (existing > 0)
                return false; // already closed

            await db.ExecuteAsync(
                "INSERT INTO Lots_Closed (LotNo, UserName, MachineName) VALUES (@LotNo, @UserName, @MachineName)",
                new { LotNo = lotNo, UserName = userName, MachineName = machineName });

            return true;
        }

        // ─────────────────────────────────────────────────────────────
        // MANUAL PTC NO.
        // Mirrors cmdUpdateManualPTCNo_Click: UPSERT VendRcvdDetail_MoreDetails
        // ─────────────────────────────────────────────────────────────
        public async Task<bool> UpdateManualPTCNoAsync(string lotNo, string manualPTCNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            int existing = await db.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM VendRcvdDetail_MoreDetails WHERE LotNo = @LotNo",
                new { LotNo = lotNo });

            if (existing > 0)
            {
                await db.ExecuteAsync(
                    "UPDATE VendRcvdDetail_MoreDetails SET ManualPTCNo = @ManualPTCNo WHERE LotNo = @LotNo",
                    new { LotNo = lotNo, ManualPTCNo = manualPTCNo });
            }
            else
            {
                await db.ExecuteAsync(
                    "INSERT INTO VendRcvdDetail_MoreDetails (LotNo, ManualPTCNo) VALUES (@LotNo, @ManualPTCNo)",
                    new { LotNo = lotNo, ManualPTCNo = manualPTCNo });
            }

            return true;
        }

        // ─────────────────────────────────────────────────────────────
        // CHECK ISSUANCE EXISTS
        // Mirrors check in mnuDeleteRcving_Click
        // ─────────────────────────────────────────────────────────────
        public async Task<int> CheckIssuanceExistsAsync(long vrdEntryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            return await db.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM VendIssdDetail WHERE Rcvd_RefID = @VRD_EntryID",
                new { VRD_EntryID = vrdEntryId });
        }

        // ─────────────────────────────────────────────────────────────
        // USER RIGHTS
        // ─────────────────────────────────────────────────────────────
        public async Task<bool> GetUserRightAsync(string rightName, string userName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            int count = await db.ExecuteScalarAsync<int>(
                @"SELECT COUNT(*) FROM UserRights
                  WHERE RightName = @RightName AND UserName = @UserName AND Value = 1",
                new { RightName = rightName, UserName = userName });
            return count > 0;
        }

        // ─────────────────────────────────────────────────────────────
        // SPLIT / TRANSFER LOT LOOKUPS
        // ─────────────────────────────────────────────────────────────
        public async Task<List<LookupItemString>> GetDistinctCustomerCodesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT DISTINCT CustCode AS Id, CustCode AS Name FROM FCustomerOrders WHERE CustCode IS NOT NULL AND CustCode <> '' ORDER BY CustCode";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<OrderLookupItem>> GetOrdersForCustomerAsync(string custCode, string? itemCode = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (!string.IsNullOrEmpty(itemCode))
            {
                const string sql = @"
                    SELECT DISTINCT FOrderItems.OrderNo, ISNULL(FCustomerOrders.InternalRefNo, '') AS InternalRefNo
                    FROM FOrderItems
                    INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo = FOrderItems.OrderNo
                    WHERE FCustomerOrders.CustCode = @CustCode
                      AND FOrderItems.CompItemCode = @ItemCode
                      AND (FCustomerOrders.OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList) OR FOrderItems.OrderNo = 'Stock-Order')
                    ORDER BY FOrderItems.OrderNo";
                return (await db.QueryAsync<OrderLookupItem>(sql, new { CustCode = custCode, ItemCode = itemCode })).ToList();
            }
            else
            {
                const string sql = @"
                    SELECT DISTINCT FOrderItems.OrderNo, ISNULL(FCustomerOrders.InternalRefNo, '') AS InternalRefNo
                    FROM FOrderItems
                    INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo = FOrderItems.OrderNo
                    WHERE FCustomerOrders.CustCode = @CustCode
                      AND FCustomerOrders.OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList)
                    ORDER BY FOrderItems.OrderNo";
                return (await db.QueryAsync<OrderLookupItem>(sql, new { CustCode = custCode })).ToList();
            }
        }

        public async Task<List<LookupItemString>> GetArticlesForOrderAsync(string orderNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT Items.ItemID AS Id,
                       '{' + VrptOrders_ForProduction.CompItemID + '} ' + VrptOrders_ForProduction.ItemName AS Name
                FROM VrptOrders_ForProduction
                INNER JOIN Items ON VrptOrders_ForProduction.CompItemID = Items.ItemID
                WHERE VrptOrders_ForProduction.OrderNo = @OrderNo
                ORDER BY VrptOrders_ForProduction.ItemName";
            return (await db.QueryAsync<LookupItemString>(sql, new { OrderNo = orderNo })).ToList();
        }

        public async Task<List<StoreLookupItem>> GetStoresAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT EntryID, StoreName FROM Stores ORDER BY StoreName";
            return (await db.QueryAsync<StoreLookupItem>(sql)).ToList();
        }

        public async Task<List<ShelfLookupItem>> GetShelvesByStoreAsync(int storeRefId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT EntryID, RackNo, ShelfNo
                FROM VStoreShelfs
                WHERE Store_RefID = @StoreRefID
                ORDER BY StoreName, RackNo, ShelfNo";
            return (await db.QueryAsync<ShelfLookupItem>(sql, new { StoreRefID = storeRefId })).ToList();
        }

        public async Task<string> GetShelfRemarksAsync(string itemCode, int processId, int shelfRefId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT TOP 1 Remarks
                FROM StockOrderOpening
                WHERE ItemID = @ItemCode AND ProcessID = @ProcessID AND Shelf_RefID = @ShelfRefID";
            return await db.ExecuteScalarAsync<string>(sql, new { ItemCode = itemCode, ProcessID = processId, ShelfRefID = shelfRefId }) ?? string.Empty;
        }

        // ─────────────────────────────────────────────────────────────
        // SPLIT / TRANSFER LOT TRANSACTIONS
        // ─────────────────────────────────────────────────────────────
        private async Task<string> GenerateNextLotNoAsync(IDbConnection db, IDbTransaction tx)
        {
            string yearPrefix = DateTime.Today.ToString("yy");
            string sql = @"
                SELECT MAX(SeqNo) FROM (
                    SELECT CAST(SUBSTRING(LotNo, 3, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT) AS SeqNo
                    FROM Lots_List WHERE LEFT(LotNo, 2) = @YearPrefix AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1
                    UNION ALL
                    SELECT CAST(SUBSTRING(LotNo, 3, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT) AS SeqNo
                    FROM VendRcvdDetail WHERE LEFT(LotNo, 2) = @YearPrefix AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1
                    UNION ALL
                    SELECT CAST(SUBSTRING(LotNo, 3, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT) AS SeqNo
                    FROM VendIssdDetail WHERE LEFT(LotNo, 2) = @YearPrefix AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1
                    UNION ALL
                    SELECT CAST(SUBSTRING(LotNo_Manual, 3, CASE WHEN CHARINDEX('-', LotNo_Manual) > 0 THEN CHARINDEX('-', LotNo_Manual) - 3 ELSE LEN(LotNo_Manual) END) AS INT) AS SeqNo
                    FROM StockOrderOpening WHERE LEFT(LotNo_Manual, 2) = @YearPrefix AND LEN(LotNo_Manual) = 7 AND ISNUMERIC(LotNo_Manual) = 1
                ) T";

            int maxSeq = await db.ExecuteScalarAsync<int?>(sql, new { YearPrefix = yearPrefix }, tx) ?? 0;
            if (maxSeq == 0)
            {
                maxSeq = 53000;
            }
            else
            {
                maxSeq++;
            }

            return $"{yearPrefix}{maxSeq:D5}";
        }

        public async Task<bool> ChangeOrderNoAsync(ChangeOrderNoRequest request)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                string toOrderNo = request.ToOrderNo;
                string toItemCode = request.ToItemCode;

                if (request.TransferToStockOrder)
                {
                    toOrderNo = "Stock-Order";
                    toItemCode = request.ToItemCode;

                    int custCount = await db.ExecuteScalarAsync<int>(
                        "SELECT COUNT(*) FROM FCustomerCatalog WHERE Custcode = 'Stock' AND Country = 'PK' AND CompItemID = @ItemCode",
                        new { ItemCode = toItemCode }, tx);
                    if (custCount == 0)
                    {
                        await db.ExecuteAsync(
                            "INSERT INTO FCustomerCatalog (Custcode, Country, ItemID, CompItemID) VALUES ('Stock', 'PK', @ItemCode, @ItemCode)",
                            new { ItemCode = toItemCode }, tx);
                    }

                    int orderItemCount = await db.ExecuteScalarAsync<int>(
                        "SELECT COUNT(*) FROM FOrderItems WHERE OrderNo = 'Stock-Order' AND CompItemCode = @ItemCode",
                        new { ItemCode = toItemCode }, tx);
                    if (orderItemCount == 0)
                    {
                        await db.ExecuteAsync(
                            @"INSERT INTO FOrderItems (OrderNo, ItemCode, CompItemCode, Qty, DeliveryDT, Stamps, Quality)
                              VALUES ('Stock-Order', @ItemCode, @ItemCode, @Qty, @Date, '', '')",
                            new { ItemCode = toItemCode, Qty = request.OriginalQty, Date = DateTime.Today }, tx);
                    }
                    else
                    {
                        await db.ExecuteAsync(
                            "UPDATE FOrderItems SET Qty = Qty + @Qty WHERE OrderNo = 'Stock-Order' AND ItemCode = @ItemCode",
                            new { ItemCode = toItemCode, Qty = request.OriginalQty }, tx);
                    }
                }

                await db.ExecuteAsync(
                    "UPDATE VendRcvdDetail SET OrderNo = @ToOrderNo, ItemCode = @ToItemCode WHERE LotNo = @LotNo",
                    new { ToOrderNo = toOrderNo, ToItemCode = toItemCode, LotNo = request.LotNo }, tx);

                await db.ExecuteAsync(
                    "UPDATE VendIssdDetail SET OrderNo = @ToOrderNo, ItemCode = @ToItemCode WHERE LotNo = @LotNo",
                    new { ToOrderNo = toOrderNo, ToItemCode = toItemCode, LotNo = request.LotNo }, tx);

                await db.ExecuteAsync(@"
                    INSERT INTO LotTransferDetails (VRD_From_RefID, VRD_To_RefID, FromOrderNo, ToOrderNo, Qty, SplitQty, Type, FromItemCode, ToItemCode)
                    VALUES (@EntryID, @EntryID, @FromOrderNo, @ToOrderNo, @Qty, @Qty, 0, @ToItemCode, @ToItemCode)",
                    new
                    {
                        EntryID = request.EntryID,
                        FromOrderNo = request.FromOrderNo,
                        ToOrderNo = toOrderNo,
                        Qty = (int)request.OriginalQty,
                        ToItemCode = toItemCode
                    }, tx);

                tx.Commit();
                return true;
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }

        public async Task<string> SplitLotAsync(SplitLotRequest request)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                string newLotNo = await GenerateNextLotNoAsync(db, tx);

                string forgeBatchNo = await db.ExecuteScalarAsync<string>(
                    "SELECT RcvID FROM VLot_With_ForgeBatchNo WHERE LotNo = @LotNo",
                    new { LotNo = request.OriginalLotNo }, tx) ?? string.Empty;

                int billCount = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM MakerPostedBillsDetail_Receivings WHERE VRD_RefID = @EntryID",
                    new { EntryID = request.EntryID }, tx);
                int notAvailableForBilling = billCount > 0 ? 1 : 0;

                await db.ExecuteAsync(
                    "UPDATE VendRcvdDetail SET RcvdQty = RcvdQty - @SplitQty WHERE EntryID = @EntryID",
                    new { SplitQty = request.SplitQty, EntryID = request.EntryID }, tx);

                await db.ExecuteAsync(@"
                    INSERT INTO VendReceived (VendID, DT, RecieptID, UserID, ProcessID, Issuance_RefID, EmpID)
                    SELECT VendID, DT, '', UserID, ProcessID, Issuance_RefID, EmpID
                    FROM VendReceived WHERE EntryID = @VR_EntryID",
                    new { VR_EntryID = request.VR_EntryID }, tx);

                long newVrRefId = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM VendReceived", transaction: tx);

                await db.ExecuteAsync(@"
                    INSERT INTO VendRcvdDetail (
                        RefID, ItemCode, RecieptID, RcvdQty, Wastage, IssQty, Rate, LotNo, ReqAuth,
                        NextProcessID, LostQty, OrderNo, CountedBy, Issue_RefID, ProcessID, RcvdWeight,
                        Opening_RefID, ReworkLot, Repair_RefID, Not_Available_For_Billing
                    )
                    SELECT
                        @NewVrRefId, @SplitItemCode, '', @SplitQty, 0, 0, Rate, @NewLotNo, ReqAuth,
                        NextProcessID, 0, @SplitOrderNo, '', Issue_RefID, ProcessID, 0,
                        Opening_RefID, ReWorkLot, Repair_RefID, @NotAvailableForBilling
                    FROM VendRcvdDetail WHERE EntryID = @EntryID",
                    new
                    {
                        NewVrRefId = newVrRefId,
                        SplitItemCode = request.SplitItemCode,
                        SplitQty = request.SplitQty,
                        NewLotNo = newLotNo,
                        SplitOrderNo = request.SplitOrderNo,
                        NotAvailableForBilling = notAvailableForBilling,
                        EntryID = request.EntryID
                    }, tx);

                long newVrdEntryId = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM VendRcvdDetail", transaction: tx);

                await db.ExecuteAsync(@"
                    INSERT INTO Lots_List (LotNo, ItemID, Lot_Type, Reference_LotNo, Batch_No, Mill_Certificate_No, Forge_Batch_No)
                    SELECT @NewLotNo, @SplitItemCode, 4, @OriginalLotNo, Batch_No, Mill_Certificate_No, @ForgeBatchNo
                    FROM Lots_List WHERE LotNo = @OriginalLotNo",
                    new
                    {
                        NewLotNo = newLotNo,
                        SplitItemCode = request.SplitItemCode,
                        OriginalLotNo = request.OriginalLotNo,
                        ForgeBatchNo = forgeBatchNo
                    }, tx);

                await db.ExecuteAsync(@"
                    INSERT INTO LotTransferDetails (VRD_From_RefID, VRD_To_RefID, FromOrderNo, ToOrderNo, Qty, SplitQty, Type, LotTransferRemarks)
                    VALUES (@FromEntryID, @NewEntryID, @FromOrderNo, @ToOrderNo, @Qty, @SplitQty, 1, @Remarks)",
                    new
                    {
                        FromEntryID = request.EntryID,
                        NewEntryID = newVrdEntryId,
                        FromOrderNo = request.FromOrderNo,
                        ToOrderNo = request.SplitOrderNo,
                        Qty = (int)request.OriginalQty,
                        SplitQty = (int)request.SplitQty,
                        Remarks = request.Remarks
                    }, tx);

                tx.Commit();
                return newLotNo;
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }

        public async Task<bool> TransferToSFStockAsync(TransferSFStockRequest request)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                await db.ExecuteAsync(@"
                    INSERT INTO StockOrderOpening (ItemID, ProcessID, Qty, Location, UserName, MachineName, Shelf_RefID, Remarks, LotNo_Manual)
                    VALUES (@ItemCode, @ProcessID, @Qty, @Location, @UserName, @MachineName, @ShelfRefID, @Remarks, @LotNo)",
                    new
                    {
                        ItemCode = request.ItemCode,
                        ProcessID = request.ProcessID,
                        Qty = (int)request.TransferQty,
                        Location = request.LocationText,
                        UserName = request.UserName,
                        MachineName = request.MachineName,
                        ShelfRefID = request.ShelfRefID,
                        Remarks = request.Remarks,
                        LotNo = request.LotNo
                    }, tx);

                long sooEntryId = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM StockOrderOpening", transaction: tx);

                int custCount = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM ForeignCustomers WHERE CustCode = 'Stock' AND Country = 'PK'", transaction: tx);
                if (custCount == 0)
                {
                    await db.ExecuteAsync("INSERT INTO ForeignCustomers (Custcode, Country) VALUES ('Stock', 'PK')", transaction: tx);
                }

                int orderCount = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM FCustomerOrders WHERE OrderNo = 'Stock-Order'", transaction: tx);
                if (orderCount == 0)
                {
                    await db.ExecuteAsync("INSERT INTO FCustomerOrders (CustCode, Country, OrderNo, DT) VALUES ('Stock', 'PK', 'Stock-Order', @Date)",
                        new { Date = DateTime.Today }, tx);
                }

                int catalogCount = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM FCustomerCatalog WHERE Custcode = 'Stock' AND Country = 'PK' AND CompItemID = @ItemCode",
                    new { ItemCode = request.ItemCode }, tx);
                if (catalogCount == 0)
                {
                    await db.ExecuteAsync("INSERT INTO FCustomerCatalog (Custcode, Country, ItemID, CompItemID) VALUES ('Stock', 'PK', @ItemCode, @ItemCode)",
                        new { ItemCode = request.ItemCode }, tx);
                }

                int orderItemCount = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM FOrderItems WHERE OrderNo = 'Stock-Order' AND CompItemCode = @ItemCode",
                    new { ItemCode = request.ItemCode }, tx);
                if (orderItemCount == 0)
                {
                    await db.ExecuteAsync(
                        @"INSERT INTO FOrderItems (OrderNo, ItemCode, CompItemCode, Qty, DeliveryDT, Stamps, Quality)
                          VALUES ('Stock-Order', @ItemCode, @ItemCode, @Qty, @Date, '', '')",
                        new { ItemCode = request.ItemCode, Qty = (int)request.TransferQty, Date = DateTime.Today }, tx);
                }
                else
                {
                    await db.ExecuteAsync(
                        "UPDATE FOrderItems SET Qty = Qty + @Qty WHERE OrderNo = 'Stock-Order' AND ItemCode = @ItemCode",
                        new { ItemCode = request.ItemCode, Qty = (int)request.TransferQty }, tx);
                }

                await db.ExecuteAsync(@"
                    INSERT INTO LotTransferDetails (VRD_From_RefID, VRD_To_RefID, FromOrderNo, ToOrderNo, Qty, SplitQty, Type, SOO_RefID)
                    VALUES (@FromEntryID, 0, @FromOrderNo, 'Stock-Order', @Qty, @SplitQty, 2, @SOO_EntryID)",
                    new
                    {
                        FromEntryID = request.EntryID,
                        FromOrderNo = request.FromOrderNo,
                        Qty = (int)request.OriginalQty,
                        SplitQty = (int)request.TransferQty,
                        SOO_EntryID = sooEntryId
                    }, tx);

                await db.ExecuteAsync(
                    "UPDATE VendRcvdDetail SET IssQty = IssQty + @TransferQty WHERE EntryID = @EntryID",
                    new { TransferQty = request.TransferQty, EntryID = request.EntryID }, tx);

                tx.Commit();
                return true;
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // INSPECTION DATA
        // ─────────────────────────────────────────────────────────────
        public async Task<InspectionDataDto> GetInspectionDataAsync(long vrdEntryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string sqlHeader = @"
                SELECT VRD.EntryID AS VRD_RefID, VRD.LotNo, VRD.ItemCode, VRD.OrderNo,
                       VRD.ProcessID, VRD.RcvdQty, VR.DT AS ReceivingDT,
                       I.ItemName, I.TipSize, I.ItemSize, I.SizeUnit,
                       P.Description AS ProcessDescription
                FROM VendRcvdDetail VRD
                INNER JOIN VendReceived VR ON VR.EntryID = VRD.RefID
                LEFT JOIN Items I ON I.ItemID = VRD.ItemCode
                LEFT JOIN Processes P ON P.ProcessID = VRD.ProcessID
                WHERE VRD.EntryID = @VRD_EntryID";

            var dto = await db.QueryFirstOrDefaultAsync<InspectionDataDto>(sqlHeader, new { VRD_EntryID = vrdEntryId });
            if (dto == null) return new InspectionDataDto();

            const string sqlProcesses = @"
                SELECT ProcessID, Code, Description
                FROM VItemProcesses
                WHERE IsExist = @ItemCode AND ProcessID = @ProcessID
                ORDER BY Description";
            dto.Processes = (await db.QueryAsync<ProcessOptionItem>(sqlProcesses, new { ItemCode = dto.ItemCode, ProcessID = dto.ProcessID })).ToList();

            const string sqlExisting = @"
                SELECT EntryID, Disposation, Comments, DT AS InspectionDT, LotStatus
                FROM VendRcvdDetailInspection
                WHERE VRD_RefID = @VRD_RefID";
            var existing = await db.QueryFirstOrDefaultAsync<dynamic>(sqlExisting, new { VRD_RefID = vrdEntryId });

            if (existing != null)
            {
                dto.EntryID = (long)existing.EntryID;
                dto.Disposation = (string)(existing.Disposation ?? string.Empty);
                dto.Comments = (string)(existing.Comments ?? string.Empty);
                dto.InspectionDT = (DateTime)(existing.InspectionDT ?? DateTime.Today);
                dto.LotStatus = (bool)(existing.LotStatus ?? false);

                const string sqlDetails = @"
                    SELECT PIP.EntryID, PIP.ParameterName, VRDID.AQL, VRDID.SampleSize,
                           VRDID.RejectOn, VRDID.TestSpecificationNo, VRDID.ActualRejection, VRDID.Status
                    FROM VendRcvdDetailInspectionDetail VRDID
                    LEFT JOIN ProcessInspectionParameters PIP ON PIP.EntryID = VRDID.PIP_RefID
                    WHERE VRDID.VRDI_RefID = @VRDI_RefID";
                dto.Parameters = (await db.QueryAsync<InspectionParameterItem>(sqlDetails, new { VRDI_RefID = dto.EntryID })).ToList();
            }

            if (!dto.Parameters.Any())
            {
                const string sqlParams = @"
                    SELECT EntryID, ParameterName, AQL, SampleSize, RejectOn,
                           TechSpecNo AS TestSpecificationNo, '0' AS ActualRejection, 'OK' AS Status
                    FROM ProcessInspectionParameters
                    WHERE ProcessID = @ProcessID
                    ORDER BY EntryID";
                dto.Parameters = (await db.QueryAsync<InspectionParameterItem>(sqlParams, new { ProcessID = dto.ProcessID })).ToList();
            }

            const string sqlTemper = @"
                SELECT TempValue
                FROM VendRcvdDetailTemperValues
                WHERE VRD_RefID = @VRD_RefID
                ORDER BY SNo, EntryID";
            dto.TemperValues = (await db.QueryAsync<string>(sqlTemper, new { VRD_RefID = vrdEntryId }))
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();

            return dto;
        }

        public async Task<List<InspectionParameterItem>> GetProcessInspectionParametersAsync(int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT EntryID, ParameterName, AQL, SampleSize, RejectOn,
                       TechSpecNo AS TestSpecificationNo, '0' AS ActualRejection, 'OK' AS Status
                FROM ProcessInspectionParameters
                WHERE ProcessID = @ProcessID
                ORDER BY EntryID";
            return (await db.QueryAsync<InspectionParameterItem>(sql, new { ProcessID = processId })).ToList();
        }

        public async Task<bool> SaveInspectionDataAsync(SaveInspectionRequest request)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                await db.ExecuteAsync(@"
                    DELETE FROM VendRcvdDetailInspectionDetail
                    WHERE VRDI_RefID IN (SELECT EntryID FROM VendRcvdDetailInspection WHERE VRD_RefID = @VRD_RefID)",
                    new { VRD_RefID = request.VRD_RefID }, tx);

                await db.ExecuteAsync(@"
                    DELETE FROM VendRcvdDetailInspection WHERE VRD_RefID = @VRD_RefID",
                    new { VRD_RefID = request.VRD_RefID }, tx);

                await db.ExecuteAsync(@"
                    INSERT INTO VendRcvdDetailInspection (VRD_RefID, Disposation, Comments, DT, LotStatus, UserName, MachineName)
                    VALUES (@VRD_RefID, @Disposation, @Comments, @DT, @LotStatus, @UserName, @MachineName)",
                    new
                    {
                        VRD_RefID = request.VRD_RefID,
                        Disposation = request.Disposation,
                        Comments = request.Comments,
                        DT = request.DT,
                        LotStatus = request.LotStatus,
                        UserName = request.UserName,
                        MachineName = request.MachineName
                    }, tx);

                long vrdiEntryId = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM VendRcvdDetailInspection", transaction: tx);

                if (request.Parameters != null && request.Parameters.Any())
                {
                    foreach (var p in request.Parameters)
                    {
                        await db.ExecuteAsync(@"
                            INSERT INTO VendRcvdDetailInspectionDetail (VRDI_RefID, PIP_RefID, AQL, SampleSize, RejectOn, TestSpecificationNo, ActualRejection, Status)
                            VALUES (@VRDI_RefID, @PIP_RefID, @AQL, @SampleSize, @RejectOn, @TestSpecificationNo, @ActualRejection, @Status)",
                            new
                            {
                                VRDI_RefID = vrdiEntryId,
                                PIP_RefID = p.EntryID,
                                AQL = p.AQL,
                                SampleSize = p.SampleSize,
                                RejectOn = p.RejectOn,
                                TestSpecificationNo = p.TestSpecificationNo,
                                ActualRejection = p.ActualRejection,
                                Status = p.Status
                            }, tx);
                    }
                }

                await db.ExecuteAsync("DELETE FROM VendRcvdDetailTemperValues WHERE VRD_RefID = @VRD_RefID",
                    new { VRD_RefID = request.VRD_RefID }, tx);

                if (request.TemperValues != null && request.TemperValues.Any())
                {
                    int sno = 1;
                    foreach (var val in request.TemperValues)
                    {
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            await db.ExecuteAsync(@"
                                INSERT INTO VendRcvdDetailTemperValues (VRD_RefID, SNo, TempValue)
                                VALUES (@VRD_RefID, @SNo, @TempValue)",
                                new { VRD_RefID = request.VRD_RefID, SNo = sno++, TempValue = val.Trim() }, tx);
                        }
                    }
                }

                tx.Commit();
                return true;
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }
    }
}


