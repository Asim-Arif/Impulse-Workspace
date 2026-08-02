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
    public class MakerPOListDataAccess : IMakerPOListDataAccess
    {
        private readonly IConfiguration _config;

        public MakerPOListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<(List<MakerPOListItem> Items, string ReportSql)> GetListAsync(MakerPOListFilter filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            DynamicParameters parameters = new DynamicParameters();

            string whereClause = string.Empty;

            // STEP 1: Due Days Mode
            if (filter.DueDaysMode > 0)
            {
                int dueDays = filter.DueDaysMode switch
                {
                    1 => 3,
                    2 => 7,
                    3 => filter.DueDaysCustom,
                    _ => 0
                };

                DateTime targetDt = DateTime.Today.AddDays(dueDays);
                DateTime dueFrom = targetDt < DateTime.Today ? targetDt : DateTime.Today;
                DateTime dueTo = targetDt < DateTime.Today ? DateTime.Today : targetDt;

                parameters.Add("@DueFrom", dueFrom);
                parameters.Add("@DueTo", dueTo);

                whereClause = @" WHERE (VVendIssued.EntryID IN(
                                    SELECT RefID FROM VendIssdDetail 
                                    INNER JOIN VendIssdDetail_ReturnDTs ON VendIssdDetail.EntryID=VendIssdDetail_ReturnDTs.VIS_RefID 
                                    WHERE VendIssdDetail_ReturnDTs.ReturnDT BETWEEN @DueFrom AND @DueTo 
                                      AND VendIssdDetail_ReturnDTs.IssQty > VendIssdDetail_ReturnDTs.RcvQty))";
            }
            else
            {
                // STEP 2: Date Range
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
                    case 5: // Custom
                        break;
                }

                parameters.Add("@DtFrom", dtFrom);
                parameters.Add("@DtTo", dtTo);

                if (!string.IsNullOrWhiteSpace(filter.LotNo) || !string.IsNullOrWhiteSpace(filter.OrderNo) || !string.IsNullOrWhiteSpace(filter.MasterPONo))
                {
                    whereClause = " WHERE OnlyDT BETWEEN '1901-01-01' AND '2090-01-01'";
                }
                else
                {
                    whereClause = " WHERE OnlyDT BETWEEN @DtFrom AND @DtTo";
                }
            }

            // STEP 3: Maker filter
            if (filter.MakerIds != null && filter.MakerIds.Any())
            {
                whereClause += $" AND VendID IN ({string.Join(",", filter.MakerIds)})";
            }

            // STEP 4: Item filter
            if (!string.IsNullOrWhiteSpace(filter.ItemId) && filter.ItemId != "0")
            {
                parameters.Add("@ItemId", filter.ItemId);
                whereClause += " AND VVendIssued.ItemID = @ItemId";
            }

            // STEP 5: Item Category filter
            if (filter.ItemCatIds != null && filter.ItemCatIds.Any())
            {
                var catEscaped = filter.ItemCatIds.Select(c => $"'{c.Replace("'", "''")}'");
                whereClause += $" AND CatID IN ({string.Join(",", catEscaped)})";
            }

            // STEP 6: Item Group filter
            if (filter.ItemGroupIds != null && filter.ItemGroupIds.Any())
            {
                whereClause += $" AND GroupID IN ({string.Join(",", filter.ItemGroupIds)})";
            }

            // STEP 7: Process filter
            if (filter.ProcessIds != null && filter.ProcessIds.Any())
            {
                whereClause += $" AND ProcessID IN ({string.Join(",", filter.ProcessIds)})";
            }

            // STEP 8: LotNo OVERRIDE (Replaces entire where clause per legacy logic)
            if (!string.IsNullOrWhiteSpace(filter.LotNo))
            {
                parameters.Add("@LotNo", filter.LotNo.Trim());
                whereClause = " WHERE LotNo = @LotNo";
            }

            // STEP 9: OrderNo
            if (!string.IsNullOrWhiteSpace(filter.OrderNo))
            {
                parameters.Add("@OrderNo", filter.OrderNo.Trim());
                whereClause += " AND VVendIssued.EntryID IN(SELECT RefID FROM VendIssdDetail WHERE OrderNo = @OrderNo)";
            }

            // STEP 10: MasterPONo
            if (!string.IsNullOrWhiteSpace(filter.MasterPONo))
            {
                parameters.Add("@MasterPONo", filter.MasterPONo.Trim());
                whereClause += " AND MasterPONo = @MasterPONo";
            }

            // STEP 11: ShowMasterPOOnly
            if (filter.ShowMasterPOOnly)
            {
                whereClause += " AND VVendIssued.MasterPONo IS NOT NULL";
            }

            // STEP 12: MasterPOOpen
            if (filter.MasterPOOpen)
            {
                whereClause += " AND DetailEntryID IN(SELECT VID_EntryID FROM VItemOpenPOs)";
            }

            // STEP 13: RepairLots
            if (filter.RepairLots)
            {
                whereClause += " AND ReWorkLot = 1";
            }

            // STEP 14: RegularLotsOnly
            if (filter.RegularLotsOnly)
            {
                whereClause += " AND ReWorkLot = 0";
            }

            // STEP 15: Bookmarks
            if (filter.Bookmarks)
            {
                whereClause += " AND BookMarkEntryID IS NOT NULL";
            }

            // STEP 16: OpenLotsOnly
            if (filter.OpenLotsOnly)
            {
                whereClause += @" AND (LotNo<>'0' 
                                  AND VVendIssued.EntryID NOT IN(SELECT Issuance_RefID FROM VendReceived) 
                                  AND LotNo IN(SELECT LotNo FROM VRunningLots_Simple) 
                                  AND LotNo NOT IN(SELECT LotNo FROM Lots_Closed))";
            }

            // STEP 17: Customers
            if (filter.CustomerCodes != null && filter.CustomerCodes.Any())
            {
                var custEscaped = filter.CustomerCodes.Select(c => $"'{c.Replace("'", "''")}'");
                whereClause += $" AND CustCode IN ({string.Join(",", custEscaped)})";
            }

            // STEP 18: RepairIssuance
            if (filter.RepairIssuance)
            {
                whereClause += " AND VVendIssued.DetailEntryID IN(SELECT VID_RefID FROM VendReWorkIssuanceDetail)";
            }

            // STEP 19: Purchaser
            if (!string.IsNullOrWhiteSpace(filter.PurchaserEmpId) && filter.PurchaserEmpId != "0")
            {
                parameters.Add("@PurchaserEmpId", filter.PurchaserEmpId);
                whereClause += " AND VVendIssued.EntryID IN(SELECT RefID FROM VendIssdDetail WHERE VID_EmpID = @PurchaserEmpId)";
            }

            string sql = $@"SELECT VVendIssued.*, VItems_Complaints.ItemID AS ComplaintItemID
                            FROM VVendIssued
                            LEFT OUTER JOIN VItems_Complaints ON VVendIssued.ItemID = VItems_Complaints.ItemID
                            {whereClause}
                            ORDER BY DT DESC";

            string reportSql = $"SELECT VVendIssued.* FROM VVendIssued LEFT OUTER JOIN VItemOpenPOs ON VVendIssued.DetailEntryID=VItemOpenPOs.VID_EntryID {whereClause} ORDER BY DT";

            var items = (await db.QueryAsync<MakerPOListItem>(sql, parameters)).ToList();
            return (items, reportSql);
        }

        public async Task<List<LookupItemInt>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = "SELECT VendID AS Id, ISNULL('{ ' + VendID1 + ' } ' + VenderName, VenderName) AS Name FROM Makers ORDER BY VenderName";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = "SELECT CustCode AS Id, CustCode AS Name FROM ForeignCustomers ORDER BY CustCode";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetItemCategoriesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = "SELECT CatID AS Id, Description AS Name FROM ItemCatagories ORDER BY Description";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<LookupItemInt>> GetItemGroupsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = "SELECT ID AS Id, Description AS Name FROM ItemGroups ORDER BY Description";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetItemsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = "SELECT ItemID AS Id, ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }' AS Name FROM Items WHERE InActive = 0 ORDER BY ItemName";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<LookupItemInt>> GetProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = "SELECT ProcessID AS Id, Description AS Name FROM Processes ORDER BY Description";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetEmployeesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = "SELECT EmpID AS Id, '{' + EmpID + '} ' + Name AS Name FROM Employees ORDER BY EmpID";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<bool> AuthorizeIssuancesAsync(IEnumerable<long> entryIds, string userName, string machineName)
        {
            if (entryIds == null || !entryIds.Any()) return true;

            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"UPDATE VendIssued 
                           SET Authorized = 1, AuthUserName = @UserName, AuthMachineName = @MachineName, AuthEntryDT = GETDATE() 
                           WHERE EntryID IN @EntryIds";

            int rows = await db.ExecuteAsync(sql, new { EntryIds = entryIds, UserName = userName, MachineName = machineName });
            return rows > 0;
        }

        public async Task<bool> CloseMakerPOAsync(long entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = "UPDATE VendIssued SET Closed = 1 WHERE EntryID = @EntryID";
            int rows = await db.ExecuteAsync(sql, new { EntryID = entryId });
            return rows > 0;
        }

        public async Task<(bool ShortLoan, bool LongLoan)> CheckLoanExistsAsync(string masterPoNo)
        {
            if (string.IsNullOrWhiteSpace(masterPoNo)) return (false, false);

            using IDbConnection db = new SqlConnection(ConnectionString);
            int shortCount = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM MakerAdvancesShort WHERE MasterPONo = @MasterPONo", new { MasterPONo = masterPoNo });
            int longCount = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM MakerAdvances WHERE MasterPONo = @MasterPONo", new { MasterPONo = masterPoNo });
            return (shortCount > 0, longCount > 0);
        }

        public async Task<int> CheckReceivingExistsAsync(long entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            return await db.ExecuteScalarAsync<int>("SELECT COUNT(Issuance_RefID) FROM VendReceived WHERE Issuance_RefID = @EntryID", new { EntryID = entryId });
        }

        public async Task<bool> DeleteIssuanceAsync(long entryId)
        {
            using SqlConnection con = new SqlConnection(ConnectionString);
            await con.OpenAsync();
            using SqlTransaction trans = con.BeginTransaction();

            try
            {
                var detailInfo = await con.QueryFirstOrDefaultAsync<(long DetailEntryID, long Rcvd_RefID, bool ReWorkLot)>(
                    "SELECT EntryID AS DetailEntryID, ISNULL(Rcvd_RefID, 0) AS Rcvd_RefID, ISNULL(ReWorkLot, 0) AS ReWorkLot FROM VendIssdDetail WHERE RefID = @RefID",
                    new { RefID = entryId }, transaction: trans);

                if (detailInfo.ReWorkLot && detailInfo.Rcvd_RefID == 0)
                {
                    string updateReworkSql = @"UPDATE VendRcvdDetailReWorkDetail 
                                               SET VendRcvdDetailReWorkDetail.IssQty = VendRcvdDetailReWorkDetail.IssQty - VendReWorkIssuanceDetail.Qty 
                                               FROM VendRcvdDetailReWorkDetail 
                                               INNER JOIN VendReWorkIssuanceDetail ON VendRcvdDetailReWorkDetail.EntryID = VRDRWD_RefID 
                                               INNER JOIN VendIssdDetail ON VendReWorkIssuanceDetail.VID_RefID = VendIssdDetail.EntryID 
                                               WHERE VendIssdDetail.EntryID = @DetailEntryID";
                    await con.ExecuteAsync(updateReworkSql, new { DetailEntryID = detailInfo.DetailEntryID }, transaction: trans);

                    await con.ExecuteAsync("DELETE FROM VendReWorkIssuanceDetail WHERE VID_RefID = @DetailEntryID",
                        new { DetailEntryID = detailInfo.DetailEntryID }, transaction: trans);
                }

                // Update SF Issuance entries
                string updateSfSql = @"UPDATE VendRcvdDetail 
                                       SET VendRcvdDetail.IssQty = VendRcvdDetail.IssQty - ISNULL(T1.QtyIssd, 0) 
                                       FROM VendRcvdDetail 
                                       INNER JOIN (SELECT SOO_RefID, SUM(QtyIssd) AS QtyIssd FROM StockOrderOpening_Issuance WHERE VID_RefID = @DetailEntryID GROUP BY SOO_RefID) T1 
                                       ON VendRcvdDetail.Opening_RefID = T1.SOO_RefID";
                await con.ExecuteAsync(updateSfSql, new { DetailEntryID = detailInfo.DetailEntryID }, transaction: trans);

                // Update standard receiving issuance entries
                string updateStandardSql = @"UPDATE VendRcvdDetail 
                                            SET VendRcvdDetail.IssQty = VendRcvdDetail.IssQty - VendIssdDetail.IssQty 
                                            FROM VendRcvdDetail 
                                            INNER JOIN VendIssdDetail ON VendRcvdDetail.EntryID = VendIssdDetail.Rcvd_RefID 
                                            WHERE VendIssdDetail.RefID = @RefID";
                await con.ExecuteAsync(updateStandardSql, new { RefID = entryId }, transaction: trans);

                // Delete from VendIssued
                await con.ExecuteAsync("DELETE FROM VendIssued WHERE EntryID = @RefID", new { RefID = entryId }, transaction: trans);

                trans.Commit();
                return true;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<bool> GetUserRightAsync(string rightName, string userName)
        {
            if (string.IsNullOrWhiteSpace(userName)) return false;

            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = $"SELECT ISNULL({rightName}, 0) FROM Users WHERE UserName = @UserName";
            try
            {
                return await db.ExecuteScalarAsync<bool>(sql, new { UserName = userName });
            }
            catch
            {
                return false;
            }
        }
    }
}
