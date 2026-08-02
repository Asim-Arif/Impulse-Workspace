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
    public class ReceiveLotDataAccess : IReceiveLotDataAccess
    {
        private readonly IConfiguration _config;
        private readonly IProductionSequenceDataAccess _sequenceDataAccess;

        public ReceiveLotDataAccess(IConfiguration config, IProductionSequenceDataAccess sequenceDataAccess)
        {
            _config = config;
            _sequenceDataAccess = sequenceDataAccess;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<LotSearchResultModel?> SearchLotAsync(string lotNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT TOP 1 VI.EntryID, VID.LotNo, ISNULL(VI.RecieptID, '') AS RecieptID, VI.VendID,
                                  ISNULL(M.VenderName, 'N/A') AS VendorName, ISNULL(M.VendID1, '') AS VendID1,
                                  VI.ProcessID, ISNULL(P.Description, 'N/A') AS ProcessName, ISNULL(VI.ItemID, '') AS ItemID, VI.DT,
                                  VI.Authorized, ISNULL(VI.Closed, 0) AS Closed, ISNULL(VI.IssEmpID, '') AS IssEmpID,
                                  CASE WHEN M.VenderName LIKE '%FACTORY%' OR M.VendID1 LIKE '%FAC%' OR M.VendID = 79 OR M.VendID = 129 OR M.VendID = (SELECT TOP 1 CAST(DataValue AS BIGINT) FROM GeneralData WHERE DataName = 'FactoryMaker') THEN 1 ELSE 0 END AS IsFactoryMaker,
                                  CASE WHEN (SELECT COUNT(VRD.EntryID) FROM VendRcvdDetail VRD WHERE VRD.Issue_RefID = VID.EntryID AND VRD.LotNo = VID.LotNo) > 0 THEN 1 ELSE 0 END AS AlreadyReceived
                           FROM VendIssdDetail VID
                           INNER JOIN VendIssued VI ON VI.EntryID = VID.RefID
                           LEFT JOIN VMakers M ON VI.VendID = M.VendID
                           LEFT JOIN Processes P ON VI.ProcessID = P.ProcessID
                           WHERE VID.EntryID = (SELECT MAX(EntryID) FROM VendIssdDetail WHERE LotNo = @LotNo)";

            return await db.QueryFirstOrDefaultAsync<LotSearchResultModel>(sql, new { LotNo = lotNo.Trim() });
        }

        public async Task<List<MasterPOLineItemModel>> GetLotLinesAsync(string lotNo, long vendIssuedEntryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VID.EntryID AS VendIssdDetailEntryID, ISNULL(VID.OrderNo, '') AS OrderNo, VID.ItemCode, ISNULL(I.ItemName, '') AS ItemName,
                                  (SELECT TOP 1 ItemPic FROM Items WHERE ItemID = VID.ItemCode) AS ItemPic,
                                  VID.IssQty, ISNULL(SUM(VRD.RcvdQty), 0) AS PrevRcvdQty, VID.Rate, ISNULL(VID.LotNo, '') AS LotNo,
                                  ISNULL(VID.ReWorkLot, 0) AS ReWorkLot, ISNULL(VID.Repair_RefID, 0) AS RepairType, VID.RcvProcessID AS ReturnProcessID, VID.Priority
                           FROM VendIssdDetail VID
                           LEFT JOIN Items I ON VID.ItemCode = I.ItemID
                           LEFT JOIN VendRcvdDetail VRD ON VID.EntryID = VRD.Issue_RefID
                           WHERE VID.RefID = @VendIssuedEntryId AND VID.LotNo = @LotNo
                           GROUP BY VID.EntryID, VID.OrderNo, VID.ItemCode, I.ItemName, VID.IssQty, VID.Rate, VID.LotNo, VID.ReWorkLot, VID.Repair_RefID, VID.RcvProcessID, VID.Priority
                           ORDER BY VID.Priority, VID.EntryID";

            return (await db.QueryAsync<MasterPOLineItemModel>(sql, new { LotNo = lotNo.Trim(), VendIssuedEntryId = vendIssuedEntryId })).ToList();
        }

        public async Task<long> SaveLotReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                // 1. Insert Header into VendReceived
                string insertHeaderSql = @"INSERT INTO VendReceived (
                                                VendID, DT, RecieptID, UserID, ProcessID, Issuance_RefID, OverTime, UserName, MachineName, TemperValue
                                           ) VALUES (
                                                @VendID, GETDATE(), '', @UserID, @ProcessID, @IssuanceRefID, @OverTime, @UserName, @MachineName, @TemperValue
                                           );
                                           SELECT SCOPE_IDENTITY();";

                long headerEntryId = await db.ExecuteScalarAsync<long>(insertHeaderSql, new
                {
                    header.VendID,
                    UserID = userId,
                    header.ProcessID,
                    header.IssuanceRefID,
                    OverTime = header.OverTime ? 1 : 0,
                    UserName = userName,
                    MachineName = machineName,
                    header.TemperValue
                }, trans);

                // 2. Insert Factory Employees
                if (header.FactoryEmpIDs != null && header.FactoryEmpIDs.Any())
                {
                    string insertEmpSql = @"INSERT INTO VendReceived_Employees (VR_RefID, EmpID) VALUES (@VR_RefID, @EmpID)";
                    foreach (var empId in header.FactoryEmpIDs)
                    {
                        await db.ExecuteAsync(insertEmpSql, new { VR_RefID = headerEntryId, EmpID = empId }, trans);
                    }
                }

                // 3. Process Lines
                foreach (var line in lines)
                {
                    string receivingReceiptId = await _sequenceDataAccess.GetNextReceivingReceiptIDAsync(header.DT);

                    // Calculate NextProcessID using legacy SP_GetNextProcID (supports RepairTypeProcesses)
                    var nextProcParams = new DynamicParameters();
                    nextProcParams.Add("@ItemCode", line.ItemCode);
                    nextProcParams.Add("@ProcID", line.ProcessID);
                    nextProcParams.Add("@ReWorkLot", line.ReWorkLot);
                    nextProcParams.Add("@RepairType", line.RepairType);
                    nextProcParams.Add("@NextProcID", dbType: DbType.Int32, direction: ParameterDirection.Output);

                    await db.ExecuteAsync("SP_GetNextProcID", nextProcParams, trans, commandType: CommandType.StoredProcedure);
                    int nextProcessId = nextProcParams.Get<int?>("@NextProcID") ?? 0;

                    // Generate LotNo if missing or update MillCertNo on existing lot
                    string lineLotNo = line.LotNo;
                    if (string.IsNullOrWhiteSpace(lineLotNo) || lineLotNo == "0")
                    {
                        lineLotNo = await _sequenceDataAccess.GetNextMainLotNoAsync(header.DT);

                        // Insert into Lots_List
                        string insertLotSql = @"INSERT INTO Lots_List (LotNo, ItemID, Lot_Type, Reference_LotNo, Batch_No, Mill_Certificate_No)
                                                VALUES (@LotNo, @ItemID, 0, '', '', @MillCertNo)";

                        await db.ExecuteAsync(insertLotSql, new { LotNo = lineLotNo, ItemID = line.ItemCode, MillCertNo = header.MillCertNo ?? "" }, trans);
                    }
                    else if (!string.IsNullOrWhiteSpace(header.MillCertNo))
                    {
                        string updateLotSql = @"UPDATE Lots_List SET Mill_Certificate_NO = @MillCertNo WHERE LotNo = @LotNo AND (Mill_Certificate_NO IS NULL OR Mill_Certificate_NO = '')";
                        await db.ExecuteAsync(updateLotSql, new { LotNo = lineLotNo, MillCertNo = header.MillCertNo }, trans);
                    }

                    // Check if authorization is required for this receiving process
                    string checkAuthSql = @"SELECT ISNULL(AuthRequired, 0) FROM Processes WHERE ProcessID = @ProcessID";
                    bool reqAuth = await db.ExecuteScalarAsync<bool>(checkAuthSql, new { line.ProcessID }, trans);

                    // Insert into VendRcvdDetail
                    string insertLineSql = @"INSERT INTO VendRcvdDetail (
                                                RefID, RecieptID, ItemCode, NextProcessID, RcvdQty, IssQty, Wastage, Rate,
                                                LotNo, ReqAuth, OrderNo, CountedBy, Issue_RefID, ProcessID, RcvdWeight, ReWorkLot, Repair_RefID
                                             ) VALUES (
                                                @RefID, @RecieptID, @ItemCode, @NextProcessID, @RcvdQty, 0, 0, @Rate,
                                                @LotNo, @ReqAuth, @OrderNo, @CountedBy, @IssueRefID, @ProcessID, 0, @ReWorkLot, @RepairType
                                             );";

                    await db.ExecuteAsync(insertLineSql, new
                    {
                        RefID = headerEntryId,
                        RecieptID = receivingReceiptId,
                        ItemCode = line.ItemCode,
                        NextProcessID = nextProcessId,
                        RcvdQty = line.RcvdQty,
                        Rate = line.Rate,
                        LotNo = lineLotNo,
                        ReqAuth = reqAuth ? 1 : 0,
                        OrderNo = line.OrderNo,
                        CountedBy = line.CountedBy ?? "",
                        IssueRefID = line.VendIssdDetailEntryID,
                        line.ProcessID,
                        ReWorkLot = line.ReWorkLot,
                        RepairType = line.RepairType
                    }, trans);

                    // Update VendIssdDetail only if authorization is NOT required
                    if (!reqAuth)
                    {
                        string updateVidSql = @"UPDATE VendIssdDetail 
                                                SET RcvdQty = ISNULL(RcvdQty, 0) + @RcvdQty 
                                                WHERE EntryID = @IssueRefID";

                        await db.ExecuteAsync(updateVidSql, new { line.RcvdQty, IssueRefID = line.VendIssdDetailEntryID }, trans);
                    }

                    // Multi-Process Auto Issuance / Receiving check
                    await ProcessMultipleIssueReceiveAsync(db, trans, line.VendIssdDetailEntryID, userName, machineName);

                    // Insert PrintSession record
                    string insertPrintSessionSql = @"INSERT INTO PrintSession (RecieptNo) VALUES (@RecieptNo)";
                    await db.ExecuteAsync(insertPrintSessionSql, new { RecieptNo = receivingReceiptId }, trans);
                }

                trans.Commit();
                return headerEntryId;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        private async Task ProcessMultipleIssueReceiveAsync(IDbConnection db, IDbTransaction trans, long issueRefId, string userName, string machineName)
        {
            string multiSql = @"SELECT VID_Multi.EntryID, VID_Multi.ProcessID, VID_Multi.Rate, VI.VendID, VI.DT,
                                       VID.ItemCode, VID.IssQty, VID.OrderNo, VID.LotNo, VID.ReWorkLot, VID.Repair_RefID
                                FROM VID_Multiple_Process_Issuance VID_Multi
                                INNER JOIN VendIssdDetail VID ON VID_Multi.VID_RefID = VID.EntryID
                                INNER JOIN VendIssued VI ON VI.EntryID = VID.RefID
                                WHERE VID_Multi.VID_RefID = @VID_RefID
                                ORDER BY VID_Multi.EntryID";

            var multiRecords = (await db.QueryAsync(multiSql, new { VID_RefID = issueRefId }, trans)).ToList();
            if (!multiRecords.Any()) return;

            int count = 0;
            int total = multiRecords.Count;

            foreach (var item in multiRecords)
            {
                count++;
                long vendId = Convert.ToInt64(item.VendID);
                DateTime dt = Convert.ToDateTime(item.DT);
                int processId = Convert.ToInt32(item.ProcessID);
                string itemCode = Convert.ToString(item.ItemCode) ?? "";
                string orderNo = Convert.ToString(item.OrderNo) ?? "";
                decimal rate = Convert.ToDecimal(item.Rate);
                decimal qty = Convert.ToDecimal(item.IssQty);
                string lotNo = Convert.ToString(item.LotNo) ?? "";
                int reWorkLot = Convert.ToInt32(item.ReWorkLot);
                int repairRefId = Convert.ToInt32(item.Repair_RefID);

                string makerIssNo = await _sequenceDataAccess.GetNextMasterPONoAsync(dt);

                string insertViSql = @"INSERT INTO VendIssued (VendID, DT, RecieptID, UserID, ProcessID, ItemID, UserName, MachineName, SpecialInstructions, ExcessQtyPercentage, MaximumRcvingsAgainstPO, MasterPONo, Authorized, IssEmpID, SteelType_RefID)
                                       VALUES (@VendID, @DT, @RecieptID, 1, @ProcessID, @ItemID, @UserName, @MachineName, 'Auto', 0, 0, '', 1, '', 0);
                                       SELECT SCOPE_IDENTITY();";

                long newViId = await db.ExecuteScalarAsync<long>(insertViSql, new
                {
                    VendID = vendId,
                    DT = dt,
                    RecieptID = makerIssNo,
                    ProcessID = processId,
                    ItemID = itemCode,
                    UserName = userName,
                    MachineName = machineName
                }, trans);

                string lastVrdSql = @"SELECT ISNULL(MAX(EntryID), 0) FROM VendRcvdDetail WHERE LotNo = @LotNo";
                long lastVrdId = await db.ExecuteScalarAsync<long>(lastVrdSql, new { LotNo = lotNo }, trans);

                string insertVidSql = @"INSERT INTO VendIssdDetail (RefID, RecieptID, ItemCode, Rate, IssQty, RcvdQty, ReqAuth, OrderNo, RcvProcessID, ReturnDT, Priority, Rcvd_RefID, LotNo, ReWorkLot, Repair_RefID)
                                        VALUES (@RefID, @RecieptID, @ItemCode, @Rate, @IssQty, @RcvdQty, 0, @OrderNo, @ProcessID, @ReturnDT, 0, @Rcvd_RefID, @LotNo, @ReWorkLot, @RepairRefID);
                                        SELECT SCOPE_IDENTITY();";

                long newVidId = await db.ExecuteScalarAsync<long>(insertVidSql, new
                {
                    RefID = newViId,
                    RecieptID = makerIssNo,
                    ItemCode = itemCode,
                    Rate = rate,
                    IssQty = qty,
                    RcvdQty = qty,
                    OrderNo = orderNo,
                    ProcessID = processId,
                    ReturnDT = dt,
                    Rcvd_RefID = lastVrdId,
                    LotNo = lotNo,
                    ReWorkLot = reWorkLot,
                    RepairRefID = repairRefId
                }, trans);

                string checkAssSql = @"SELECT COUNT(*) FROM VendAssItems WHERE VendID = @VendID AND ProcessID = @ProcessID AND ItemID = @ItemID";
                int assCount = await db.ExecuteScalarAsync<int>(checkAssSql, new { VendID = vendId, ProcessID = processId, ItemID = itemCode }, trans);
                if (assCount == 0)
                {
                    string insertAssSql = @"INSERT INTO VendAssItems (VendID, ProcessID, ItemID, Rate, Unit, Remarks) VALUES (@VendID, @ProcessID, @ItemID, 1, 'Pcs', 'Auto-Multi')";
                    await db.ExecuteAsync(insertAssSql, new { VendID = vendId, ProcessID = processId, ItemID = itemCode }, trans);
                }

                string insertVrSql = @"INSERT INTO VendReceived (VendID, DT, RecieptID, UserID, ProcessID, Issuance_RefID, OverTime, UserName, MachineName, TemperValue)
                                       VALUES (@VendID, @DT, '', 1, @ProcessID, @IssuanceRefID, 0, @UserName, @MachineName, '');
                                       SELECT SCOPE_IDENTITY();";

                long newVrId = await db.ExecuteScalarAsync<long>(insertVrSql, new
                {
                    VendID = vendId,
                    DT = dt,
                    ProcessID = processId,
                    IssuanceRefID = newViId,
                    UserName = userName,
                    MachineName = machineName
                }, trans);

                string snoProcSql = @"SELECT ISNULL(SNo, 0) FROM ItemProcesses WHERE ProcessID = @ProcessID AND ItemID = @ItemID";
                int snoProc = await db.ExecuteScalarAsync<int?>(snoProcSql, new { ProcessID = processId, ItemID = itemCode }, trans) ?? 0;

                string nextProcSql = @"SELECT TOP 1 ProcessID FROM ItemProcesses WHERE ItemID = @ItemID AND SNo > @SNo ORDER BY SNo";
                int nextProcId = await db.ExecuteScalarAsync<int?>(nextProcSql, new { ItemID = itemCode, SNo = snoProc }, trans) ?? 0;

                decimal vrdIssdQty = count == total ? 0 : qty;

                string insertVrdSql = @"INSERT INTO VendRcvdDetail (RefID, RecieptID, ItemCode, NextProcessID, RcvdQty, IssQty, Wastage, Rate, LotNo, ReqAuth, OrderNo, CountedBy, Issue_RefID, ProcessID, RcvdWeight, ReWorkLot, Repair_RefID)
                                        VALUES (@RefID, '', @ItemCode, @NextProcessID, @RcvdQty, @IssQty, 0, @Rate, @LotNo, 0, @OrderNo, '', @Issue_RefID, @ProcessID, 0, @ReWorkLot, @RepairRefID);";

                await db.ExecuteAsync(insertVrdSql, new
                {
                    RefID = newVrId,
                    ItemCode = itemCode,
                    NextProcessID = nextProcId,
                    RcvdQty = qty,
                    IssQty = vrdIssdQty,
                    Rate = rate,
                    LotNo = lotNo,
                    OrderNo = orderNo,
                    Issue_RefID = newVidId,
                    ProcessID = processId,
                    ReWorkLot = reWorkLot,
                    RepairRefID = repairRefId
                }, trans);
            }
        }
    }
}
