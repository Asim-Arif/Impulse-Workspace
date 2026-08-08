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
    public class CreateDispatchListDataAccess : ICreateDispatchListDataAccess
    {
        private readonly IConfiguration _config;

        public CreateDispatchListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("ImpulseConnection")
            ?? string.Empty;

        public async Task<List<LookupItemString>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT DISTINCT CustCode AS Id, CustCode AS Name FROM ForeignCustomers ORDER BY CustCode";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<string>> GetCountriesByCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE CustCode = @CustCode ORDER BY Country";
            return (await db.QueryAsync<string>(sql, new { CustCode = custCode })).ToList();
        }

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

        public async Task<List<FinalizedLotItem>> GetAvailableFinalizedLotsAsync(DispatchListFilter filter)
        {
            if (string.IsNullOrWhiteSpace(filter.CustCode))
            {
                return new List<FinalizedLotItem>();
            }

            using IDbConnection db = new SqlConnection(ConnectionString);

            string sql = @"
                SELECT VRD_EntryID, VendID1, VenderName, InternalRefNo, DT,
                       Description, ProcessID, ItemCode, ItemName, ItemSize, SizeUnit,
                       RcvdQty, ISNULL(Wastage, 0) AS Wastage, ISNULL(LostQty, 0) AS LostQty,
                       ISNULL(IssQty, 0) AS IssQty, ISNULL(ReWorkQty, 0) AS ReWorkQty,
                       OrderNo, LotNo
                FROM VVendReceivingList
                WHERE OnlyDT >= @DtFrom AND OnlyDT <= @DtTo
                  AND CustCode = @CustCode
                  AND (RcvdQty - ISNULL(Wastage,0) - ISNULL(LostQty,0) - ISNULL(ReWorkQty,0)) > IssQty
                  AND ReqAuth = 0
                  AND NextProcessID IS NULL";

            var p = new DynamicParameters();
            p.Add("@CustCode", filter.CustCode);
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

            if (!string.IsNullOrWhiteSpace(filter.OrderNo))
            {
                sql += " AND OrderNo LIKE @OrderNo";
                p.Add("@OrderNo", $"%{filter.OrderNo.Trim()}%");
            }

            if (!string.IsNullOrWhiteSpace(filter.LotNo))
            {
                sql += " AND LotNo LIKE @LotNo";
                p.Add("@LotNo", $"%{filter.LotNo.Trim()}%");
            }

            sql += " ORDER BY DT";

            var list = (await db.QueryAsync<FinalizedLotItem>(sql, p)).ToList();
            foreach (var item in list)
            {
                item.DispatchQty = item.AvailQty;
            }

            return list;
        }

        public async Task<string?> GetInnerLabelReportNameAsync(string custCode, string country, bool otherLabel = false, bool manualQty = false)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            if (manualQty)
            {
                string col = otherLabel ? "Inner_Label_Manual_II" : "Inner_Label_Manual_I";
                string sql = $"SELECT TOP 1 {col} FROM ForeignCustomers WHERE CustCode = @CustCode AND Country = @Country";
                string? reportName = await db.ExecuteScalarAsync<string>(sql, new { CustCode = custCode, Country = country });
                return string.IsNullOrWhiteSpace(reportName) ? null : reportName.Trim();
            }
            else
            {
                const string sql = "SELECT TOP 1 InnerPackingLabel FROM ForeignCustomers WHERE CustCode = @CustCode AND Country = @Country";
                string? reportName = await db.ExecuteScalarAsync<string>(sql, new { CustCode = custCode, Country = country });
                if (string.IsNullOrWhiteSpace(reportName)) return null;

                reportName = reportName.Trim();
                if (otherLabel)
                {
                    if (reportName.EndsWith(".rpt", StringComparison.OrdinalIgnoreCase))
                    {
                        reportName = reportName.Substring(0, reportName.Length - 4) + "_II.rpt";
                    }
                    else
                    {
                        reportName = reportName + "_II.rpt";
                    }
                }
                return reportName;
            }
        }

        public async Task<LoadedDispatchListForEdit?> GetDispatchListForEditAsync(long entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string headerSql = @"
                SELECT EntryID, DispatchListNo, CustCode, Country, ISNULL(Finalyzed, 0) AS Finalyzed
                FROM DispatchList
                WHERE EntryID = @EntryID";

            var header = await db.QueryFirstOrDefaultAsync<LoadedDispatchListForEdit>(headerSql, new { EntryID = entryId });
            if (header == null) return null;

            const string linesSql = @"
                SELECT DispatchListDetail_Inners.EntryID AS InnerEntryID,
                       DispatchListDetail_Inners.Qty,
                       DispatchListDetails_Adv.ItemCode,
                       Items.ItemName,
                       DispatchListDetails_Adv.ProcessID,
                       Processes.Description AS ProcessDescription,
                       ISNULL(DispatchListDetail_Inners.CartonNo, 0) AS CartonNo,
                       ISNULL(DispatchListDetail_Inners.Lots, '') AS LotNo,
                       ISNULL(DispatchListDetail_Inners.InternalNos, '') AS OrderNo
                FROM DispatchListDetails_Adv
                INNER JOIN VDispatchListDetail_Inners DispatchListDetail_Inners ON DispatchListDetails_Adv.EntryID = DispatchListDetail_Inners.RefID
                INNER JOIN Items ON DispatchListDetails_Adv.ItemCode = Items.ItemID
                LEFT JOIN Processes ON DispatchListDetails_Adv.ProcessID = Processes.ProcessID
                WHERE DispatchListDetails_Adv.RefID = @EntryID
                ORDER BY DispatchListDetails_Adv.EntryID, DispatchListDetail_Inners.EntryID";

            header.Items = (await db.QueryAsync<DispatchListStagedItem>(linesSql, new { EntryID = entryId })).ToList();
            return header;
        }

        public async Task<SaveDispatchListResult> SaveDispatchListAsync(SaveDispatchListRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.DispatchListNo))
            {
                return new SaveDispatchListResult { Success = false, Message = "Please provide a Dispatch List Number." };
            }

            if (request.Items == null || !request.Items.Any())
            {
                return new SaveDispatchListResult { Success = false, Message = "No items staged for dispatch." };
            }

            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                long lEntryID = request.EntryID;

                if (lEntryID == 0)
                {
                    // 1. Insert New DispatchList Header
                    await db.ExecuteAsync(@"
                        INSERT INTO DispatchList (DispatchListNo, UserName, MachineName, DT, EntryDT, CustCode, Country, NewFormat)
                        VALUES (@DispatchListNo, @UserName, @MachineName, @DT, @EntryDT, @CustCode, @Country, 1)",
                        new
                        {
                            DispatchListNo = request.DispatchListNo,
                            UserName = request.UserName,
                            MachineName = request.MachineName,
                            DT = DateTime.Today,
                            EntryDT = DateTime.Now,
                            CustCode = request.CustCode,
                            Country = request.Country
                        }, tx);

                    lEntryID = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM DispatchList", transaction: tx);
                }
                else
                {
                    // Update Header
                    await db.ExecuteAsync(@"
                        UPDATE DispatchList
                        SET DispatchListNo = @DispatchListNo,
                            CustCode = @CustCode,
                            Country = @Country
                        WHERE EntryID = @EntryID",
                        new
                        {
                            DispatchListNo = request.DispatchListNo,
                            CustCode = request.CustCode,
                            Country = request.Country,
                            EntryID = lEntryID
                        }, tx);

                    // Handle Deleted Rows: revert IssQty and remove inner records
                    if (request.DeletedInnerIDs != null && request.DeletedInnerIDs.Any())
                    {
                        var delDetails = (await db.QueryAsync<(long VRD_RefID, int Qty)>(@"
                            SELECT VRD_RefID, Qty
                            FROM DispatchListDetail_VRD
                            WHERE DLDC_RefID IN @DelIDs",
                            new { DelIDs = request.DeletedInnerIDs }, tx)).ToList();

                        foreach (var d in delDetails)
                        {
                            await db.ExecuteAsync(@"
                                UPDATE VendRcvdDetail
                                SET IssQty = IssQty - @Qty
                                WHERE EntryID = @VRD_RefID",
                                new { Qty = d.Qty, VRD_RefID = d.VRD_RefID }, tx);
                        }

                        await db.ExecuteAsync("DELETE FROM DispatchListDetail_VRD WHERE DLDC_RefID IN @DelIDs", new { DelIDs = request.DeletedInnerIDs }, tx);
                        await db.ExecuteAsync("DELETE FROM DispatchListDetail_Inners WHERE EntryID IN @DelIDs", new { DelIDs = request.DeletedInnerIDs }, tx);
                        await db.ExecuteAsync("DELETE FROM DispatchListDetails_Adv WHERE RefID = @RefID AND EntryID NOT IN (SELECT RefID FROM DispatchListDetail_Inners)", new { RefID = lEntryID }, tx);
                    }
                }

                // 2. Separate existing lines from newly added lines
                var existingLines = request.Items.Where(x => x.InnerEntryID > 0).ToList();
                var newLines = request.Items.Where(x => x.InnerEntryID == 0).ToList();

                // Update CartonNo on existing lines
                foreach (var ex in existingLines)
                {
                    await db.ExecuteAsync(@"
                        UPDATE DispatchListDetail_Inners
                        SET CartonNo = @CartonNo
                        WHERE EntryID = @InnerEntryID",
                        new { CartonNo = ex.CartonNo, InnerEntryID = ex.InnerEntryID }, tx);
                }

                // Insert New Lines
                if (newLines.Any())
                {
                    var groups = newLines.GroupBy(x => new { x.ItemCode, x.ProcessID }).ToList();
                    var groupDetailMap = new Dictionary<string, long>();

                    foreach (var grp in groups)
                    {
                        decimal grpQty = grp.Sum(x => x.Qty);

                        await db.ExecuteAsync(@"
                            INSERT INTO DispatchListDetails_Adv (RefID, ItemCode, Qty, ProcessID)
                            VALUES (@RefID, @ItemCode, @Qty, @ProcessID)",
                            new
                            {
                                RefID = lEntryID,
                                ItemCode = grp.Key.ItemCode,
                                Qty = (int)grpQty,
                                ProcessID = grp.Key.ProcessID
                            }, tx);

                        long lDLD_EntryID = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM DispatchListDetails_Adv", transaction: tx);
                        groupDetailMap[$"{grp.Key.ItemCode}_{grp.Key.ProcessID}"] = lDLD_EntryID;

                        await db.ExecuteAsync(@"
                            UPDATE Items
                            SET InHand = ISNULL(InHand, 0) + @Qty
                            WHERE ItemID = @ItemID",
                            new { Qty = grpQty, ItemID = grp.Key.ItemCode }, tx);
                    }

                    int maxInnerNo = await db.ExecuteScalarAsync<int?>(@"
                        SELECT MAX(InnerNo)
                        FROM DispatchListDetail_Inners
                        WHERE RefID IN (SELECT EntryID FROM DispatchListDetails_Adv WHERE RefID = @RefID)",
                        new { RefID = lEntryID }, tx) ?? 0;

                    int innerIdx = maxInnerNo + 1;
                    foreach (var item in newLines)
                    {
                        string groupKey = $"{item.ItemCode}_{item.ProcessID}";
                        long dldRefId = groupDetailMap.TryGetValue(groupKey, out var val) ? val : lEntryID;

                        await db.ExecuteAsync(@"
                            INSERT INTO DispatchListDetail_Inners (RefID, InnerNo, Qty, CartonNo)
                            VALUES (@RefID, @InnerNo, @Qty, @CartonNo)",
                            new
                            {
                                RefID = dldRefId,
                                InnerNo = innerIdx++,
                                Qty = (int)item.Qty,
                                CartonNo = item.CartonNo
                            }, tx);

                        long lDLDC_EntryID = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM DispatchListDetail_Inners", transaction: tx);

                        await db.ExecuteAsync(@"
                            INSERT INTO DispatchListDetail_VRD (DLD_RefID, DLDC_RefID, VRD_RefID, Qty)
                            VALUES (@DLD_RefID, @DLDC_RefID, @VRD_RefID, @Qty)",
                            new
                            {
                                DLD_RefID = dldRefId,
                                DLDC_RefID = lDLDC_EntryID,
                                VRD_RefID = item.VRD_EntryID,
                                Qty = (int)item.Qty
                            }, tx);

                        await db.ExecuteAsync(@"
                            UPDATE VendRcvdDetail
                            SET IssQty = IssQty + @Qty
                            WHERE EntryID = @VRD_EntryID",
                            new { Qty = item.Qty, VRD_EntryID = item.VRD_EntryID }, tx);
                    }
                }

                tx.Commit();

                return new SaveDispatchListResult
                {
                    Success = true,
                    DispatchListEntryID = lEntryID,
                    DispatchListNo = request.DispatchListNo,
                    Message = request.EntryID > 0
                        ? $"Dispatch List '{request.DispatchListNo}' updated successfully."
                        : $"Dispatch List '{request.DispatchListNo}' created successfully."
                };
            }
            catch (Exception ex)
            {
                tx.Rollback();
                return new SaveDispatchListResult
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }
    }
}
