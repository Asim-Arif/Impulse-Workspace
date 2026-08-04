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
    }
}
