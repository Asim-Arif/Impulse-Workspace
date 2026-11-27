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
    public class AuthorizeReceivedDataAccess : IAuthorizeReceivedDataAccess
    {
        private readonly IConfiguration _config;

        public AuthorizeReceivedDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<List<PendingAuthorizeReceivedItemModel>> GetPendingAuthorizeReceivedItemsAsync(AuthorizeReceivedFilterModel filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VR.EntryID, VR.RefID, VR.RecieptID, ISNULL(VR.VendID, 0) AS VendID,
                                  ISNULL(VR.VendID1, '') AS VendID1, ISNULL(VR.VenderName, '') AS VenderName,
                                  ISNULL(VR.Supervisor, '') AS Supervisor, ISNULL(VR.OrderNo, '') AS OrderNo,
                                  VR.ItemCode, ISNULL(VR.ItemDescription, '') AS ItemDescription,
                                  ISNULL(VR.LotNo, '') AS LotNo, ISNULL(VR.ProcessID, 0) AS ProcessID,
                                  ISNULL(VR.Process, '') AS Process, ISNULL(VR.RcvdQty, 0) AS RcvdQty,
                                  ISNULL(VR.Wastage, 0) AS Wastage, ISNULL(VR.LostQty, 0) AS LostQty,
                                  ISNULL(VR.ReWorkQty, 0) AS ReWorkQty, ISNULL(VR.RepairAmt, 0) AS RepairAmt,
                                  ISNULL(VR.Rate, 0) AS Rate, ISNULL(VR.Issue_RefID, 0) AS Issue_RefID,
                                  ISNULL(VR.Issuance_RefID, 0) AS Issuance_RefID, ISNULL(VR.RcvdWeight, 0) AS RcvdWeight,
                                  ISNULL(VR.Unit, 'Pcs') AS Unit, (SELECT ISNULL(Insp_EmpID, '') FROM VendRcvdDetail WHERE EntryID = VR.EntryID) AS Insp_EmpID
                           FROM VVendRcvItemsrpt VR
                           WHERE VR.ReqAuth = 1
                             AND (@RefId IS NULL OR @RefId = 0 OR VR.RefID = @RefId)
                             AND (@VendID = 0 OR VR.VendID = @VendID)
                             AND (@ProcessID = 0 OR VR.ProcessID = @ProcessID)
                             AND (@LotNo = '' OR VR.LotNo = @LotNo)
                             AND (@DateFrom IS NULL OR VR.DT >= @DateFrom)
                             AND (@DateTo IS NULL OR VR.DT <= @DateTo)
                             AND (@SearchText = '' OR VR.RecieptID LIKE '%' + @SearchText + '%' OR VR.ItemCode LIKE '%' + @SearchText + '%' OR VR.ItemDescription LIKE '%' + @SearchText + '%')
                           ORDER BY VR.EntryID";

            var items = (await db.QueryAsync<PendingAuthorizeReceivedItemModel>(sql, new
            {
                RefId = filter.RefId ?? 0,
                VendID = filter.VendID,
                ProcessID = filter.ProcessID,
                LotNo = filter.LotNo ?? "",
                DateFrom = filter.DateFrom,
                DateTo = filter.DateTo,
                SearchText = filter.SearchText ?? ""
            })).ToList();

            foreach (var item in items)
            {
                item.IsChecked = true;
            }

            return items;
        }

        public async Task<List<WastageTypeLookupModel>> GetWastageTypesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EntryID, WastageName AS Description, ISNULL(WastageType, 1) AS WastageType FROM WastageTypes WHERE ISNULL(Closed, 0) = 0 ORDER BY WastageName";
            return (await db.QueryAsync<WastageTypeLookupModel>(sql)).ToList();
        }

        public async Task<List<RepairTypeLookupModel>> GetRepairTypesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EntryID, RepairType AS Description FROM RepairTypes ORDER BY RepairType";
            return (await db.QueryAsync<RepairTypeLookupModel>(sql)).ToList();
        }

        public async Task<List<EmployeeLookupModel>> GetInspectorsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EmpID, Name, ISNULL(Designation, '') AS Designation FROM Employees WHERE ISNULL(Active, 1) = 1 ORDER BY Name";
            return (await db.QueryAsync<EmployeeLookupModel>(sql)).ToList();
        }

        public async Task<List<ProcessPOLookupModel>> GetProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessID, Description, ISNULL(SNO, 0) AS SNO FROM Processes ORDER BY SNO";
            return (await db.QueryAsync<ProcessPOLookupModel>(sql)).ToList();
        }

        public async Task<List<MakerPOLookupModel>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VendID, ISNULL(VendID1, '') AS VendID1, VenderName FROM VMakers WHERE Active = 1 ORDER BY VenderName";
            return (await db.QueryAsync<MakerPOLookupModel>(sql)).ToList();
        }

        public async Task<List<LotProcessHistoryItemModel>> GetLotProcessHistoryAsync(string lotNo, long upToVrdEntryId)
        {
            if (string.IsNullOrWhiteSpace(lotNo)) return new List<LotProcessHistoryItemModel>();

            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                IF OBJECT_ID('dbo.GetAllLots_F') IS NOT NULL
                BEGIN
                    SELECT VRD.EntryID,
                           ISNULL(VRD.LotNo, '') AS LotNo,
                           ISNULL(P.Code, '') AS ProcessCode,
                           ISNULL(P.Description, '') AS ProcessName,
                           ISNULL(VR.VendID, 0) AS VendID,
                           ISNULL(M.VenderName, '') AS VenderName
                    FROM VendRcvdDetail VRD
                    INNER JOIN VendReceived VR ON VR.EntryID = VRD.RefID
                    LEFT JOIN Processes P ON P.ProcessID = VRD.ProcessID
                    LEFT JOIN Makers M ON M.VendID = VR.VendID
                    WHERE VRD.LotNo IN (SELECT LotNo FROM dbo.GetAllLots_F(@LotNo))
                      AND VRD.EntryID <= @UpToEntryID
                    ORDER BY VRD.EntryID;
                END
                ELSE
                BEGIN
                    SELECT VRD.EntryID,
                           ISNULL(VRD.LotNo, '') AS LotNo,
                           ISNULL(P.Code, '') AS ProcessCode,
                           ISNULL(P.Description, '') AS ProcessName,
                           ISNULL(VR.VendID, 0) AS VendID,
                           ISNULL(M.VenderName, '') AS VenderName
                    FROM VendRcvdDetail VRD
                    INNER JOIN VendReceived VR ON VR.EntryID = VRD.RefID
                    LEFT JOIN Processes P ON P.ProcessID = VRD.ProcessID
                    LEFT JOIN Makers M ON M.VendID = VR.VendID
                    WHERE VRD.LotNo = @LotNo
                      AND VRD.EntryID <= @UpToEntryID
                    ORDER BY VRD.EntryID;
                END";

            return (await db.QueryAsync<LotProcessHistoryItemModel>(sql, new { LotNo = lotNo.Trim(), UpToEntryID = upToVrdEntryId })).ToList();
        }

        private static int RoundToInt(decimal value) =>
            Convert.ToInt32(Math.Round(value, 0, MidpointRounding.AwayFromZero));

        private sealed class LotProcessChargeRow
        {
            public long EntryID { get; set; }
            public long VendID { get; set; }
            public long Issue_RefID { get; set; }
            public decimal Rate { get; set; }
        }

        private sealed class ResolvedWastageEntry
        {
            public decimal Qty { get; set; }
            public int WastageType { get; set; }
            public long TargetVendID { get; set; }
            public long ReturnToVrdRefID { get; set; }
            public string TargetEmpID { get; set; } = string.Empty;
        }

        private static async Task<long> UpsertVendIssdDetailMoreDetailsAsync(IDbConnection db, IDbTransaction trans, long vendIssdDetailId, int wastageQty)
        {
            if (vendIssdDetailId <= 0 || wastageQty <= 0) return 0;

            long existingMdId = await db.QueryFirstOrDefaultAsync<long?>(
                "SELECT TOP 1 EntryID FROM VendIssdDetail_MoreDetails WHERE VID_RefID = @VID_RefID",
                new { VID_RefID = vendIssdDetailId }, trans) ?? 0;

            int issRcvBalance = 0;
            try
            {
                decimal? bal = await db.QueryFirstOrDefaultAsync<decimal?>(
                    @"SELECT CAST(ISNULL(IssQty, 0) - ISNULL(RcvdQty, 0) - ISNULL(IssRcvBalance, 0) - ISNULL(UnAuthRcvdQty, 0) + ISNULL(Wastage_Return, 0) AS DECIMAL(18, 4))
                      FROM VMakerIssItems
                      WHERE EntryID = @VID_RefID",
                    new { VID_RefID = vendIssdDetailId }, trans);
                if (bal.HasValue)
                {
                    issRcvBalance = RoundToInt(bal.Value);
                }
            }
            catch
            {
                issRcvBalance = 0;
            }

            if (existingMdId == 0)
            {
                const string insertMdSql = @"
                    INSERT INTO VendIssdDetail_MoreDetails (VID_RefID, EmpID, Wastage_Return, IssRcvBalance)
                    VALUES (@VID_RefID, '', @WastageReturn, @IssRcvBalance);
                    SELECT CAST(SCOPE_IDENTITY() AS BIGINT);";

                return await db.ExecuteScalarAsync<long>(insertMdSql, new
                {
                    VID_RefID = vendIssdDetailId,
                    WastageReturn = wastageQty,
                    IssRcvBalance = issRcvBalance
                }, trans);
            }
            else
            {
                const string updateMdSql = @"
                    UPDATE VendIssdDetail_MoreDetails
                    SET Wastage_Return = ISNULL(Wastage_Return, 0) + @WastageReturn,
                        IssRcvBalance = @IssRcvBalance
                    WHERE EntryID = @EntryID";

                await db.ExecuteAsync(updateMdSql, new
                {
                    WastageReturn = wastageQty,
                    IssRcvBalance = issRcvBalance,
                    EntryID = existingMdId
                }, trans);

                return existingMdId;
            }
        }

        public async Task<bool> SaveAuthorizationAsync(List<PendingAuthorizeReceivedItemModel> items, string inspectorEmpId, string userName, string machineName, int userId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                foreach (var item in items.Where(i => i.IsChecked))
                {
                    decimal totalWastage = item.WastageBreakdowns.Any() ? item.WastageBreakdowns.Sum(w => w.Qty) : item.Wastage;
                    decimal totalReWork = item.ReWorkBreakdowns.Any() ? item.ReWorkBreakdowns.Sum(r => r.Qty) : item.ReWorkQty;

                    // 1. Update VendRcvdDetail (atomic guard: only authorize rows still pending ReqAuth = 1)
                    string updateVrdSql = @"UPDATE VendRcvdDetail 
                                            SET ReqAuth = 0, 
                                                Wastage = @Wastage, 
                                                RepairAmt = @RepairAmt, 
                                                LostQty = 0, 
                                                ReWorkQty = @ReWorkQty, 
                                                Insp_EmpID = @InspEmpID 
                                            WHERE EntryID = @EntryID AND ReqAuth = 1";

                    int updatedRows = await db.ExecuteAsync(updateVrdSql, new
                    {
                        Wastage = totalWastage,
                        RepairAmt = item.RepairAmt,
                        ReWorkQty = totalReWork,
                        InspEmpID = inspectorEmpId ?? "",
                        EntryID = item.EntryID
                    }, trans);

                    if (updatedRows == 0)
                    {
                        throw new InvalidOperationException(
                            $"Lot #{item.LotNo} (Receipt Detail #{item.EntryID}) has already been authorized. Please refresh the list.");
                    }

                    // 2. Insert Rework Details
                    var reworkList = item.ReWorkBreakdowns?.ToList() ?? new List<ReWorkBreakdownItemModel>();
                    if (!reworkList.Any() && totalReWork > 0)
                    {
                        int defaultRepairId = await db.ExecuteScalarAsync<int>(
                            "SELECT TOP 1 EntryID FROM RepairTypes ORDER BY EntryID", transaction: trans);
                        if (defaultRepairId > 0)
                        {
                            reworkList.Add(new ReWorkBreakdownItemModel
                            {
                                VRD_RefID = item.EntryID,
                                Repair_RefID = defaultRepairId,
                                Qty = totalReWork
                            });
                        }
                    }

                    if (reworkList.Any())
                    {
                        foreach (var rw in reworkList)
                        {
                            string insertRwSql = @"INSERT INTO VendRcvdDetailReWorkDetail (VRD_RefID, Repair_RefID, Qty) 
                                                   VALUES (@VRD_RefID, @Repair_RefID, @Qty);
                                                   SELECT SCOPE_IDENTITY();";

                            long rwEntryId = await db.ExecuteScalarAsync<long>(insertRwSql, new
                            {
                                VRD_RefID = item.EntryID,
                                Repair_RefID = rw.Repair_RefID,
                                Qty = rw.Qty
                            }, trans);

                            string insertRwProcSql = @"INSERT INTO VendRcvdDetailReWorkDetailProcesses (VRDRWD_RefID, ProcessID)
                                                       SELECT @VRDRWD_RefID, ProcessID 
                                                       FROM RepairTypeProcesses 
                                                       WHERE Repair_RefID = @Repair_RefID";

                            await db.ExecuteAsync(insertRwProcSql, new { VRDRWD_RefID = rwEntryId, Repair_RefID = rw.Repair_RefID }, trans);
                        }
                    }

                    // 3. Insert Wastage Details & Resolve Charge Targets (Exact Legacy frmAuthorizeIssRcvAdv.frm Parity)
                    decimal wastageCh2Vendor = 0;
                    decimal wastageCh2Originator = 0;
                    decimal wastageCh2Selection = 0;
                    var resolvedWastageEntries = new List<ResolvedWastageEntry>();

                    var wastageList = item.WastageBreakdowns?.ToList() ?? new List<WastageBreakdownItemModel>();
                    if (!wastageList.Any() && totalWastage > 0)
                    {
                        var defaultWt = await db.QueryFirstOrDefaultAsync<WastageTypeLookupModel>(
                            "SELECT TOP 1 EntryID, WastageName AS Description, ISNULL(WastageType, 1) AS WastageType FROM WastageTypes WHERE ISNULL(Closed, 0) = 0 ORDER BY EntryID",
                            transaction: trans);
                        if (defaultWt != null && defaultWt.EntryID > 0)
                        {
                            wastageList.Add(new WastageBreakdownItemModel
                            {
                                VRD_RefID = item.EntryID,
                                Wastage_RefID = defaultWt.EntryID,
                                Qty = totalWastage,
                                WastageType = defaultWt.WastageType,
                                SelectionTargetMode = 0,
                                ReturnTo_VRD_RefID = defaultWt.WastageType == 3 ? item.EntryID : 0,
                                VendID = item.VendID
                            });
                        }
                    }

                    if (wastageList.Any())
                    {
                        foreach (var w in wastageList)
                        {
                            int resolvedWastageType = w.WastageType;
                            if (resolvedWastageType <= 0 && w.Wastage_RefID > 0)
                            {
                                resolvedWastageType = await db.QueryFirstOrDefaultAsync<int?>(
                                    "SELECT ISNULL(WastageType, 0) FROM WastageTypes WHERE EntryID = @EntryID",
                                    new { EntryID = w.Wastage_RefID }, trans) ?? 0;
                            }

                            long targetVendId = 0;
                            long returnToVrdRefId = 0;
                            string targetEmpId = string.Empty;

                            if (resolvedWastageType == 1)
                            {
                                wastageCh2Vendor += w.Qty;
                                targetVendId = item.VendID;
                            }
                            else if (resolvedWastageType == 2)
                            {
                                wastageCh2Originator += w.Qty;
                                const string origVendSql = @"
                                    SELECT TOP 1 ISNULL(VR.VendID, 0)
                                    FROM VendReceived VR
                                    INNER JOIN VendRcvdDetail VRD ON VR.EntryID = VRD.RefID
                                    WHERE VRD.EntryID = (SELECT MIN(EntryID) FROM VendRcvdDetail WHERE LotNo = @LotNo)";
                                targetVendId = await db.ExecuteScalarAsync<long?>(origVendSql, new { LotNo = item.LotNo }, trans) ?? item.VendID;
                            }
                            else if (resolvedWastageType == 3)
                            {
                                wastageCh2Selection += w.Qty;
                                if (w.SelectionTargetMode == 2)
                                {
                                    // Company / No Charge
                                    targetVendId = 0;
                                    returnToVrdRefId = 0;
                                    targetEmpId = string.Empty;
                                }
                                else if (w.SelectionTargetMode == 1 || (!string.IsNullOrWhiteSpace(w.EmpID) && w.VendID == 0))
                                {
                                    // Charge to Employee
                                    targetVendId = 0;
                                    returnToVrdRefId = 0;
                                    targetEmpId = w.EmpID?.Trim() ?? string.Empty;
                                }
                                else
                                {
                                    // Charge to Selected Vendor / Lot Process
                                    targetVendId = w.VendID > 0 ? w.VendID : item.VendID;
                                    returnToVrdRefId = w.ReturnTo_VRD_RefID;
                                    targetEmpId = string.Empty;
                                }
                            }

                            string insertWastageSql = @"INSERT INTO VendRcvdDetailWastageDetail (
                                                            VRD_RefID, Wastage_RefID, Qty, WastageType, ReturnTo_VRD_RefID, VendID, EmpID
                                                        ) VALUES (
                                                            @VRD_RefID, @Wastage_RefID, @Qty, @WastageType, @ReturnTo_VRD_RefID, @VendID, @EmpID
                                                        )";

                            await db.ExecuteAsync(insertWastageSql, new
                            {
                                VRD_RefID = item.EntryID,
                                Wastage_RefID = w.Wastage_RefID,
                                Qty = w.Qty,
                                WastageType = resolvedWastageType,
                                ReturnTo_VRD_RefID = returnToVrdRefId,
                                VendID = targetVendId,
                                EmpID = targetEmpId
                            }, trans);

                            resolvedWastageEntries.Add(new ResolvedWastageEntry
                            {
                                Qty = w.Qty,
                                WastageType = resolvedWastageType,
                                TargetVendID = targetVendId,
                                ReturnToVrdRefID = returnToVrdRefId,
                                TargetEmpID = targetEmpId
                            });
                        }
                    }

                    // 4. MakerRepair Ledger Insertions (All 3 Wastage Types — Legacy frmAuthorizeIssRcvAdv.frm L2246-L2386)
                    const string insertMrFullSql = @"
                        INSERT INTO MakerRepair (
                            VendID, DT, Rcvd_RefID, Qty, Rate, UserName, MachineName,
                            VendIssdDetail_RefID, RepairAmt, Wastage_Originator, ShowInBilling,
                            VRD_RefID_ChargeTo, EmpID, VRD_RefID
                        ) VALUES (
                            @VendID, CAST(GETDATE() AS DATE), @Rcvd_RefID, @Qty, @Rate, @UserName, @MachineName,
                            @VendIssdDetail_RefID, @RepairAmt, @WastageOriginator, @ShowInBilling,
                            @VRD_RefID_ChargeTo, @EmpID, @VRD_RefID
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS BIGINT);";

                    // 4A. WastageType = 1 (Charge to Current Vendor) or RepairAmt > 0
                    if ((wastageCh2Vendor > 0 || item.RepairAmt > 0) && item.EntryID > 0)
                    {
                        decimal issueRate = await db.QueryFirstOrDefaultAsync<decimal?>(
                            "SELECT CAST(ISNULL(Rate, 0) AS DECIMAL(18, 4)) FROM VendIssdDetail WHERE EntryID = @IssueRefID",
                            new { IssueRefID = item.Issue_RefID }, trans) ?? item.Rate;

                        await db.ExecuteScalarAsync<long>(insertMrFullSql, new
                        {
                            VendID = item.VendID,
                            Rcvd_RefID = item.EntryID,
                            Qty = RoundToInt(wastageCh2Vendor),
                            Rate = RoundToInt(issueRate),
                            UserName = userName,
                            MachineName = machineName,
                            VendIssdDetail_RefID = item.Issue_RefID,
                            RepairAmt = RoundToInt(item.RepairAmt),
                            WastageOriginator = false,
                            ShowInBilling = false,
                            VRD_RefID_ChargeTo = item.EntryID,
                            EmpID = string.Empty,
                            VRD_RefID = item.EntryID
                        }, trans);
                    }

                    // 4B. WastageType = 2 (Charge to Originator)
                    if (wastageCh2Originator > 0 && !string.IsNullOrWhiteSpace(item.LotNo))
                    {
                        var origRow = await db.QueryFirstOrDefaultAsync<LotProcessChargeRow>(@"
                            SELECT TOP 1
                                   VRD.EntryID,
                                   ISNULL(VR.VendID, 0) AS VendID,
                                   ISNULL(VRD.Issue_RefID, 0) AS Issue_RefID,
                                   CAST(COALESCE(NULLIF(VRD.Rate, 0), VID.Rate, 0) AS DECIMAL(18, 4)) AS Rate
                            FROM VendRcvdDetail VRD
                            INNER JOIN VendReceived VR ON VR.EntryID = VRD.RefID
                            LEFT JOIN VendIssdDetail VID ON VID.EntryID = VRD.Issue_RefID
                            WHERE VRD.LotNo = @LotNo
                            ORDER BY VRD.EntryID ASC",
                            new { LotNo = item.LotNo }, trans);

                        if (origRow != null && origRow.EntryID > 0)
                        {
                            int billedCnt = await db.ExecuteScalarAsync<int>(
                                "SELECT COUNT(VRD_RefID) FROM MakerPostedBillsDetail_Receivings WHERE VRD_RefID = @VRD_RefID",
                                new { VRD_RefID = origRow.EntryID }, trans);

                            int qtyOrigInt = RoundToInt(wastageCh2Originator);
                            long mrOrigId = await db.ExecuteScalarAsync<long>(insertMrFullSql, new
                            {
                                VendID = origRow.VendID,
                                Rcvd_RefID = item.EntryID,
                                Qty = qtyOrigInt,
                                Rate = RoundToInt(origRow.Rate),
                                UserName = userName,
                                MachineName = machineName,
                                VendIssdDetail_RefID = item.Issue_RefID,
                                RepairAmt = 0,
                                WastageOriginator = true,
                                ShowInBilling = billedCnt > 0,
                                VRD_RefID_ChargeTo = origRow.EntryID,
                                EmpID = string.Empty,
                                VRD_RefID = origRow.EntryID
                            }, trans);

                            long mdOrigId = await UpsertVendIssdDetailMoreDetailsAsync(db, trans, origRow.Issue_RefID, qtyOrigInt);
                            if (mrOrigId > 0 && mdOrigId > 0)
                            {
                                await db.ExecuteAsync(
                                    "INSERT INTO MakerRepair_VID_MD_Details (MR_RefID, VIS_MD_RefID, Qty) VALUES (@MR_RefID, @VIS_MD_RefID, @Qty)",
                                    new { MR_RefID = mrOrigId, VIS_MD_RefID = mdOrigId, Qty = qtyOrigInt }, trans);
                            }
                        }
                    }

                    // 4C. WastageType = 3 (Charge to Selected Vendor / Employee across Lot Processes up to current VRD)
                    if (wastageCh2Selection > 0 && !string.IsNullOrWhiteSpace(item.LotNo))
                    {
                        var lotProcessRows = (await db.QueryAsync<LotProcessChargeRow>(@"
                            SELECT VRD.EntryID,
                                   ISNULL(VR.VendID, 0) AS VendID,
                                   ISNULL(VRD.Issue_RefID, 0) AS Issue_RefID,
                                   CAST(COALESCE(NULLIF(VRD.Rate, 0), VID.Rate, 0) AS DECIMAL(18, 4)) AS Rate
                            FROM VendRcvdDetail VRD
                            INNER JOIN VendReceived VR ON VR.EntryID = VRD.RefID
                            LEFT JOIN VendIssdDetail VID ON VID.EntryID = VRD.Issue_RefID
                            WHERE VRD.LotNo = @LotNo
                              AND VRD.EntryID <= @CurrentVrdEntryID
                            ORDER BY VRD.EntryID ASC",
                            new { LotNo = item.LotNo, CurrentVrdEntryID = item.EntryID }, trans)).ToList();

                        foreach (var sel in resolvedWastageEntries.Where(x => x.WastageType == 3 && x.Qty > 0))
                        {
                            // Skip MakerRepair posting when explicitly marked Company / No Charge (VendID = 0 & EmpID = '')
                            if (sel.TargetVendID <= 0 && string.IsNullOrWhiteSpace(sel.TargetEmpID))
                            {
                                continue;
                            }

                            int selQtyInt = RoundToInt(sel.Qty);
                            long mdSelId = 0;

                            if (sel.ReturnToVrdRefID > 0)
                            {
                                long retIssueRefId = await db.QueryFirstOrDefaultAsync<long?>(
                                    "SELECT ISNULL(Issue_RefID, 0) FROM VendRcvdDetail WHERE EntryID = @EntryID",
                                    new { EntryID = sel.ReturnToVrdRefID }, trans) ?? 0;

                                mdSelId = await UpsertVendIssdDetailMoreDetailsAsync(db, trans, retIssueRefId, selQtyInt);
                            }

                            long matchedMrEntryId = 0;
                            foreach (var procRow in lotProcessRows)
                            {
                                int billedCnt = await db.ExecuteScalarAsync<int>(
                                    "SELECT COUNT(VRD_RefID) FROM MakerPostedBillsDetail_Receivings WHERE VRD_RefID = @VRD_RefID",
                                    new { VRD_RefID = procRow.EntryID }, trans);

                                bool showInBilling = billedCnt > 0;
                                long chargeToEntryId;

                                if (sel.TargetVendID > 0 && procRow.VendID == sel.TargetVendID)
                                {
                                    chargeToEntryId = procRow.EntryID;
                                }
                                else
                                {
                                    chargeToEntryId = 0;
                                    showInBilling = true;
                                }

                                if (!string.IsNullOrWhiteSpace(sel.TargetEmpID))
                                {
                                    chargeToEntryId = 0;
                                }

                                long mrEntryId = await db.ExecuteScalarAsync<long>(insertMrFullSql, new
                                {
                                    VendID = sel.TargetVendID,
                                    Rcvd_RefID = item.EntryID,
                                    Qty = selQtyInt,
                                    Rate = RoundToInt(procRow.Rate),
                                    UserName = userName,
                                    MachineName = machineName,
                                    VendIssdDetail_RefID = item.Issue_RefID,
                                    RepairAmt = 0,
                                    WastageOriginator = false,
                                    ShowInBilling = showInBilling,
                                    VRD_RefID_ChargeTo = chargeToEntryId,
                                    EmpID = sel.TargetEmpID,
                                    VRD_RefID = procRow.EntryID
                                }, trans);

                                if (sel.ReturnToVrdRefID > 0 && procRow.EntryID == sel.ReturnToVrdRefID)
                                {
                                    matchedMrEntryId = mrEntryId;
                                }
                            }

                            if (matchedMrEntryId > 0 && mdSelId > 0)
                            {
                                await db.ExecuteAsync(
                                    "INSERT INTO MakerRepair_VID_MD_Details (MR_RefID, VIS_MD_RefID, Qty) VALUES (@MR_RefID, @VIS_MD_RefID, @Qty)",
                                    new { MR_RefID = matchedMrEntryId, VIS_MD_RefID = mdSelId, Qty = selQtyInt }, trans);
                            }
                        }
                    }

                    // 5. Insert Audit Log
                    string insertAuthLogSql = @"INSERT INTO VendRcvAuthDetail (UserID, RcvID, DT) VALUES (@UserID, @RcvID, GETDATE())";
                    await db.ExecuteAsync(insertAuthLogSql, new { UserID = userId, RcvID = item.EntryID }, trans);

                    // 6. Update Issued Quantity (Release Lot)
                    if (item.Issue_RefID > 0)
                    {
                        string updateVidSql = @"UPDATE VendIssdDetail 
                                                SET RcvdQty = ISNULL(RcvdQty, 0) + @RcvdQty 
                                                WHERE EntryID = @IssueRefID";

                        await db.ExecuteAsync(updateVidSql, new { RcvdQty = item.RcvdQty, IssueRefID = item.Issue_RefID }, trans);
                    }
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
