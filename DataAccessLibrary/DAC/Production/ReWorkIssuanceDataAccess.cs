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
    public class ReWorkIssuanceDataAccess : IReWorkIssuanceDataAccess
    {
        private readonly IConfiguration _config;

        public ReWorkIssuanceDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("ImpulseConnection")
            ?? string.Empty;

        public async Task<List<LookupItemString>> GetArticlesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT ItemID AS Id,
                       '{ ' + ItemID + ' } ' + ItemName AS Name
                FROM Items
                WHERE InActive = 0
                ORDER BY ItemName";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<LookupItemInt>> GetRepairTypesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT EntryID AS Id, RepairType AS Name FROM RepairTypes ORDER BY RepairType";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<ReWorkAvailableItem>> GetAvailableReWorkLotsAsync(ReWorkIssuanceFilter filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            string sql = @"
                SELECT EntryID, VRD_RefID, LotNo, OrderNo, VendID, VendID1, VenderName,
                       Repair_RefID, RepairType, Qty, IssQty, Rate, ItemCode, ItemName,
                       ProcessID, Description AS ProcessDescription, DT
                FROM VVendRcvdDetailReWorkDetail
                WHERE DT >= @DtFrom AND DT <= @DtTo
                  AND ItemCode = @ArticleId
                  AND Qty > IssQty";

            var p = new DynamicParameters();
            p.Add("@ArticleId", filter.ArticleId);
            p.Add("@DtFrom", filter.DtFrom.Date);
            p.Add("@DtTo", filter.DtTo.Date.AddDays(1).AddTicks(-1));

            if (!string.IsNullOrWhiteSpace(filter.LotNo))
            {
                sql += " AND LotNo = @LotNo";
                p.Add("@LotNo", filter.LotNo.Trim());
            }

            if (filter.RepairTypeId > 0)
            {
                sql += " AND Repair_RefID = @RepairTypeId";
                p.Add("@RepairTypeId", filter.RepairTypeId);
            }

            sql += " ORDER BY EntryID";

            return (await db.QueryAsync<ReWorkAvailableItem>(sql, p)).ToList();
        }

        public async Task<List<LookupItemInt>> GetMakersForRepairProcessAsync(int repairRefId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            int processId = await db.ExecuteScalarAsync<int>(
                "SELECT TOP 1 ProcessID FROM RepairTypeProcesses WHERE Repair_RefID = @RepairRefID ORDER BY SeqNo",
                new { RepairRefID = repairRefId });

            if (processId == 0)
            {
                // Fallback to all makers
                return (await db.QueryAsync<LookupItemInt>(
                    "SELECT VendID AS Id, ISNULL('{ ' + VendID1 + ' } ' + VenderName, VenderName) AS Name FROM Makers ORDER BY VenderName")).ToList();
            }

            const string sql = @"
                SELECT VendID AS Id, ISNULL('{ ' + VendID1 + ' } ' + VenderName, VenderName) AS Name
                FROM Makers
                WHERE VendID IN (SELECT VendID FROM MakerProcesses WHERE ProcessID = @ProcessID)
                ORDER BY VenderName";

            var list = (await db.QueryAsync<LookupItemInt>(sql, new { ProcessID = processId })).ToList();
            if (!list.Any())
            {
                return (await db.QueryAsync<LookupItemInt>(
                    "SELECT VendID AS Id, ISNULL('{ ' + VendID1 + ' } ' + VenderName, VenderName) AS Name FROM Makers ORDER BY VenderName")).ToList();
            }
            return list;
        }


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

        public async Task<ReWorkIssuanceResult> SaveReWorkIssuanceAsync(SaveReWorkIssuanceRequest request)
        {
            if (request.StagedItems == null || !request.StagedItems.Any())
            {
                return new ReWorkIssuanceResult { Success = false, Message = "No items staged for issuance." };
            }

            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                var first = request.StagedItems.First();
                string strItemID = request.ItemID;
                int lVendID = request.MakerID;
                decimal dRate = first.Rate;
                string strOrderNo = first.OrderNo;
                int iRepairType = first.Repair_RefID;
                decimal totalQty = request.StagedItems.Sum(x => x.Qty);

                int lProcessID = await db.ExecuteScalarAsync<int>(
                    "SELECT TOP 1 ProcessID FROM RepairTypeProcesses WHERE Repair_RefID = @RepairRefID ORDER BY SeqNo",
                    new { RepairRefID = iRepairType }, tx);

                // 1. Generate Next Lot No
                string lLotNo = await GenerateNextLotNoAsync(db, tx);

                // 2. Fetch original Batch_No and Mill_Certificate_No from reference lot
                string strRefLotNo = first.LotNo;
                string strBatchNo = await db.ExecuteScalarAsync<string>(
                    "SELECT TOP 1 Batch_No FROM Lots_List WHERE LotNo = @LotNo",
                    new { LotNo = strRefLotNo }, tx) ?? string.Empty;
                string strMilNo = await db.ExecuteScalarAsync<string>(
                    "SELECT TOP 1 Mill_Certificate_No FROM Lots_List WHERE LotNo = @LotNo",
                    new { LotNo = strRefLotNo }, tx) ?? string.Empty;

                // 3. Insert into Lots_List (Lot_Type = 5 for Rework)
                await db.ExecuteAsync(@"
                    INSERT INTO Lots_List (LotNo, ItemID, Lot_Type, Reference_LotNo, Batch_No, Mill_Certificate_No)
                    VALUES (@LotNo, @ItemID, 5, @Reference_LotNo, @Batch_No, @Mill_Certificate_No)",
                    new
                    {
                        LotNo = lLotNo,
                        ItemID = strItemID,
                        Reference_LotNo = strRefLotNo,
                        Batch_No = strBatchNo,
                        Mill_Certificate_No = strMilNo
                    }, tx);

                // 4. Generate Maker Issuance Receipt No (M-ISU-ddMMyyN)
                string dtKey = DateTime.Today.ToString("ddMMyy");
                int lastMakerIssNo = await db.ExecuteScalarAsync<int?>(@"
                    SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 12) AS INT))
                    FROM VendIssued
                    WHERE CONVERT(VARCHAR, DT, 6) = @DtFormatted",
                    new { DtFormatted = DateTime.Today.ToString("dd MMM yy") }, tx) ?? 0;
                string strMakerIssNo = $"M-ISU-{dtKey}{lastMakerIssNo + 1}";

                // 5. Generate Detail Receipt No (ISU-ddMMyyN)
                int lastIssDetailNo = await db.ExecuteScalarAsync<int?>(@"
                    SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 10) AS INT))
                    FROM VMakerIssItems
                    WHERE CONVERT(VARCHAR, DT, 6) = @DtFormatted",
                    new { DtFormatted = DateTime.Today.ToString("dd MMM yy") }, tx) ?? 0;
                string strIssIDDetail = $"ISU-{dtKey}{lastIssDetailNo + 1}";

                // 6. Insert VendIssued Header
                await db.ExecuteAsync(@"
                    INSERT INTO VendIssued (
                        VendID, DT, RecieptID, UserID, ProcessID, ItemID,
                        UserName, MachineName, SpecialInstructions, ExcessQtyPercentage,
                        MaximumRcvingsAgainstPO, Authorized
                    )
                    VALUES (
                        @VendID, @DT, @RecieptID, 0, @ProcessID, @ItemID,
                        @UserName, @MachineName, '', 100, 100, 1
                    )",
                    new
                    {
                        VendID = lVendID,
                        DT = DateTime.Today,
                        RecieptID = strMakerIssNo,
                        ProcessID = lProcessID,
                        ItemID = strItemID,
                        UserName = request.UserName,
                        MachineName = request.MachineName
                    }, tx);

                long lRefID = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM VendIssued", transaction: tx);

                // 7. Insert VendIssdDetail Row
                await db.ExecuteAsync(@"
                    INSERT INTO VendIssdDetail (
                        RefID, RecieptID, ItemCode, Rate, IssQty, ReqAuth, OrderNo,
                        RcvProcessID, ReturnDT, Priority, Rcvd_RefID, LotNo, ReWorkLot, Repair_RefID
                    )
                    VALUES (
                        @RefID, @RecieptID, @ItemCode, @Rate, @IssQty, 0, @OrderNo,
                        @RcvProcessID, @ReturnDT, 1, 0, @LotNo, 1, @Repair_RefID
                    )",
                    new
                    {
                        RefID = lRefID,
                        RecieptID = strIssIDDetail,
                        ItemCode = strItemID,
                        Rate = dRate,
                        IssQty = totalQty,
                        OrderNo = strOrderNo,
                        RcvProcessID = lProcessID,
                        ReturnDT = DateTime.Today,
                        LotNo = lLotNo,
                        Repair_RefID = iRepairType
                    }, tx);

                long lEntryIDDetail = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM VendIssdDetail", transaction: tx);

                // 8. Insert VendIssdDetail_ReturnDTs
                await db.ExecuteAsync(@"
                    INSERT INTO VendIssdDetail_ReturnDTs (VIS_RefID, IssQty, RcvQty, ReturnDT)
                    VALUES (@VIS_RefID, @IssQty, 0, @ReturnDT)",
                    new
                    {
                        VIS_RefID = lEntryIDDetail,
                        IssQty = totalQty,
                        ReturnDT = DateTime.Today
                    }, tx);

                // 9. Update each staged rework lot & log VendReWorkIssuanceDetail
                foreach (var staged in request.StagedItems)
                {
                    await db.ExecuteAsync(@"
                        UPDATE VendRcvdDetailReWorkDetail
                        SET IssQty = IssQty + @Qty
                        WHERE EntryID = @EntryID",
                        new { Qty = staged.Qty, EntryID = staged.EntryID }, tx);

                    await db.ExecuteAsync(@"
                        INSERT INTO VendReWorkIssuanceDetail (VID_RefID, VRDRWD_RefID, Qty)
                        VALUES (@VID_RefID, @VRDRWD_RefID, @Qty)",
                        new
                        {
                            VID_RefID = lEntryIDDetail,
                            VRDRWD_RefID = staged.EntryID,
                            Qty = staged.Qty
                        }, tx);
                }

                tx.Commit();

                return new ReWorkIssuanceResult
                {
                    Success = true,
                    NewLotNo = lLotNo,
                    MakerIssNo = strMakerIssNo,
                    Message = $"Lot #{lLotNo} generated successfully."
                };
            }
            catch (Exception ex)
            {
                tx.Rollback();
                return new ReWorkIssuanceResult
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }
    }
}
