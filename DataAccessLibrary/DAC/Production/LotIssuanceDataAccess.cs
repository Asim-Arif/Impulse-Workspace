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
    public class LotIssuanceDataAccess : ILotIssuanceDataAccess
    {
        private readonly IConfiguration _config;
        private readonly IProductionSequenceDataAccess _sequenceDataAccess;

        public LotIssuanceDataAccess(IConfiguration config, IProductionSequenceDataAccess sequenceDataAccess)
        {
            _config = config;
            _sequenceDataAccess = sequenceDataAccess;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<List<ProcessPOLookupModel>> GetProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessID, Description, ISNULL(SNO, 0) AS SNO FROM Processes ORDER BY SNO";
            return (await db.QueryAsync<ProcessPOLookupModel>(sql)).ToList();
        }

        public async Task<List<MakerPOLookupModel>> GetMakersForProcessAsync(int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT M.VendID, ISNULL(M.VendID1, '') AS VendID1, M.VenderName,
                                  ISNULL(B.MakerCapacity, 0) AS MakerCapacity, ISNULL(B.Balance, 0) AS Balance
                           FROM VMakers M
                           LEFT JOIN VMakerBalances B ON M.VendID = B.VendID
                           WHERE M.Active = 1 
                             AND (M.VendID IN (SELECT VendID FROM MakerProcesses WHERE ProcessID = @ProcessID)
                                  OR M.VendID = 79 OR M.VenderName LIKE '%FACTORY%' OR M.VendID1 LIKE '%FAC%')
                           ORDER BY M.VenderName";

            return (await db.QueryAsync<MakerPOLookupModel>(sql, new { ProcessID = processId })).ToList();
        }

        public async Task<LotIssuanceLookupResultModel> LookupLotForIssuanceAsync(string lotNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string trimmedLotNo = lotNo.Trim();

            // 1. Get the latest receiving record in VendRcvdDetail for this Lot No
            string rcvdSql = @"SELECT TOP 1 VRD.EntryID AS VendIssdDetailEntryID, ISNULL(VRD.OrderNo, '') AS OrderNo, VRD.ItemCode, ISNULL(I.ItemName, '') AS ItemName,
                                      (SELECT TOP 1 ItemPic FROM Items WHERE ItemID = VRD.ItemCode) AS ItemPic,
                                      VRD.LotNo, ISNULL(L.Batch_No, '') AS BatchNo,
                                      (VRD.RcvdQty - ISNULL(VRD.IssQty, 0) - ISNULL(VRD.Wastage, 0) - ISNULL(VRD.ReWorkQty, 0) - ISNULL(VRD.LostQty, 0)) AS AvailableQty,
                                      VRD.Rate, ISNULL(I.Unit, 'Pcs') AS Unit,
                                      ISNULL(VRD.NextProcessID, 0) AS TargetProcessID, ISNULL(P.Description, '') AS TargetProcessName,
                                      ISNULL(VRD.NextProcessID, 0) AS ReturnProcessID, ISNULL(P.Description, '') AS ReturnProcessName,
                                      0 AS Priority,
                                      ISNULL(VRD.ReWorkLot, 0) AS ReWorkLot, ISNULL(VRD.Repair_RefID, 0) AS RepairType,
                                      ISNULL(VRD.ReqAuth, 0) AS ReqAuth
                               FROM VendRcvdDetail VRD
                               INNER JOIN VendReceived VR ON VRD.RefID = VR.EntryID
                               LEFT JOIN Items I ON VRD.ItemCode = I.ItemID
                               LEFT JOIN Processes P ON VRD.NextProcessID = P.ProcessID
                               LEFT JOIN Lots_List L ON VRD.LotNo = L.LotNo
                               WHERE VRD.LotNo = @LotNo
                               ORDER BY VRD.EntryID DESC";

            var latestRcvd = await db.QueryFirstOrDefaultAsync<AvailableLotIssuanceItemModel>(rcvdSql, new { LotNo = trimmedLotNo });

            if (latestRcvd != null)
            {
                // Check if receiving entry is pending authorization
                if (latestRcvd.ReqAuth == 1)
                {
                    return new LotIssuanceLookupResultModel
                    {
                        IsFound = true,
                        AlreadyIssued = true,
                        Message = $"Lot No [{trimmedLotNo}] is pending authorization and cannot be issued yet."
                    };
                }

                // Check if this latest receiving record has already been issued out (exact legacy check: COUNT(Rcvd_RefID) FROM VendIssdDetail)
                string checkIssuedSql = @"SELECT COUNT(*) FROM VendIssdDetail WHERE Rcvd_RefID = @RcvdEntryID";
                int issuedCount = await db.ExecuteScalarAsync<int>(checkIssuedSql, new { RcvdEntryID = latestRcvd.VendIssdDetailEntryID });

                if (issuedCount > 0)
                {
                    return new LotIssuanceLookupResultModel
                    {
                        IsFound = true,
                        AlreadyIssued = true,
                        Message = $"Lot No [{trimmedLotNo}] is Already Issued."
                    };
                }

                latestRcvd.IssuanceQty = latestRcvd.AvailableQty;
                return new LotIssuanceLookupResultModel
                {
                    IsFound = true,
                    AlreadyIssued = false,
                    Items = new List<AvailableLotIssuanceItemModel> { latestRcvd }
                };
            }

            // 2. If not found in VendRcvdDetail, check if lot is issued in VendIssdDetail (pending receiving)
            string checkIssdSql = @"SELECT COUNT(*) FROM VendIssdDetail WHERE LotNo = @LotNo";
            int issdCount = await db.ExecuteScalarAsync<int>(checkIssdSql, new { LotNo = trimmedLotNo });

            if (issdCount > 0)
            {
                return new LotIssuanceLookupResultModel
                {
                    IsFound = true,
                    AlreadyIssued = true,
                    Message = $"Lot No [{trimmedLotNo}] is Already Issued."
                };
            }

            return new LotIssuanceLookupResultModel
            {
                IsFound = false,
                AlreadyIssued = false,
                Message = $"No valid issuable receiving record found for Lot No: {trimmedLotNo}"
            };
        }

        public async Task<List<AvailableLotIssuanceItemModel>> GetAvailableIssuanceItemsAsync(int processId, long vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VID.EntryID AS VendIssdDetailEntryID, ISNULL(VID.OrderNo, '') AS OrderNo, VID.ItemCode, ISNULL(I.ItemName, '') AS ItemName,
                                  (SELECT TOP 1 ItemPic FROM Items WHERE ItemID = VID.ItemCode) AS ItemPic,
                                  ISNULL(VID.LotNo, '') AS LotNo, ISNULL(VID.Batch_No, '') AS BatchNo,
                                  (VID.IssQty - ISNULL(VID.RcvdQty, 0)) AS AvailableQty,
                                  VID.Rate, ISNULL(I.Unit, 'Pcs') AS Unit, ISNULL(VID.RcvProcessID, 0) AS TargetProcessID,
                                  ISNULL(P.Description, '') AS TargetProcessName, ISNULL(VID.RcvProcessID, 0) AS ReturnProcessID,
                                  ISNULL(P.Description, '') AS ReturnProcessName, ISNULL(VID.Priority, 0) AS Priority,
                                  ISNULL(VID.ReWorkLot, 0) AS ReWorkLot, ISNULL(VID.Repair_RefID, 0) AS RepairType
                           FROM VendIssdDetail VID
                           INNER JOIN VendIssued VI ON VID.RefID = VI.EntryID
                           LEFT JOIN Items I ON VID.ItemCode = I.ItemID
                           LEFT JOIN Processes P ON VID.RcvProcessID = P.ProcessID
                           WHERE (VID.IssQty - ISNULL(VID.RcvdQty, 0)) > 0
                             AND (@ProcessID = 0 OR VID.RcvProcessID = @ProcessID)
                             AND (@VendID = 0 OR VI.VendID = @VendID)
                           ORDER BY VID.EntryID DESC";

            return (await db.QueryAsync<AvailableLotIssuanceItemModel>(sql, new { ProcessID = processId, VendID = vendId })).ToList();
        }

        public async Task<long> SaveLotIssuanceAsync(CreateLotIssuanceHeaderModel header, List<CreateLotIssuanceLineModel> lines, string userName, int userId, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                string receiptId = await _sequenceDataAccess.GetNextMasterPONoAsync(header.DT);

                // 1. Insert Header into VendIssued
                string insertHeaderSql = @"INSERT INTO VendIssued (
                                                VendID, DT, RecieptID, UserID, ProcessID, ItemID, UserName, MachineName, SpecialInstructions,
                                                ExcessQtyPercentage, MaximumRcvingsAgainstPO, MasterPONo, Authorized, IssEmpID, SteelType_RefID,
                                                SampleProvided, DrawingProvided, ForgingProvided, SteelProvided
                                           ) VALUES (
                                                @VendID, @DT, @RecieptID, @UserID, @ProcessID, @ItemID, @UserName, @MachineName, @SpecialInstructions,
                                                0, 0, @MasterPONo, 1, @IssEmpID, 0,
                                                @SampleProvided, @DrawingProvided, @ForgingProvided, @SteelProvided
                                           );
                                           SELECT SCOPE_IDENTITY();";

                string mainItemCode = lines.FirstOrDefault()?.ItemCode ?? "";

                long headerEntryId = await db.ExecuteScalarAsync<long>(insertHeaderSql, new
                {
                    header.VendID,
                    header.DT,
                    RecieptID = receiptId,
                    UserID = userId,
                    header.ProcessID,
                    ItemID = mainItemCode,
                    UserName = userName,
                    MachineName = machineName,
                    SpecialInstructions = header.SpecialInstructions ?? "",
                    MasterPONo = header.MasterPONo ?? "",
                    IssEmpID = header.IssEmpID ?? "",
                    header.SampleProvided,
                    header.DrawingProvided,
                    header.ForgingProvided,
                    header.SteelProvided
                }, trans);

                // 2. Process Lines
                foreach (var line in lines)
                {
                    string subLotNo = line.LotNo;
                    if (string.IsNullOrWhiteSpace(subLotNo))
                    {
                        subLotNo = await _sequenceDataAccess.GetNextMainLotNoAsync(header.DT);
                    }

                    string insertLineSql = @"INSERT INTO VendIssdDetail (
                                                RefID, RecieptID, ItemCode, Rate, IssQty, RcvdQty, ReqAuth, OrderNo, RcvProcessID,
                                                ReturnDT, Priority, Rcvd_RefID, LotNo, ReWorkLot, Repair_RefID, Batch_No, CountedBy
                                             ) VALUES (
                                                @RefID, @RecieptID, @ItemCode, @Rate, @IssQty, 0, 0, @OrderNo, @RcvProcessID,
                                                @ReturnDT, @Priority, @IssueRefID, @LotNo, @ReWorkLot, @RepairType, @BatchNo, @CountedBy
                                             );
                                             SELECT SCOPE_IDENTITY();";

                    long lineEntryId = await db.ExecuteScalarAsync<long>(insertLineSql, new
                    {
                        RefID = headerEntryId,
                        RecieptID = receiptId,
                        ItemCode = line.ItemCode,
                        Rate = line.Rate,
                        IssQty = line.IssQty,
                        OrderNo = line.OrderNo ?? "",
                        RcvProcessID = line.RcvProcessID,
                        ReturnDT = line.ReturnDT,
                        Priority = line.Priority,
                        IssueRefID = line.IssueRefID,
                        LotNo = subLotNo,
                        ReWorkLot = line.ReWorkLot,
                        RepairType = line.RepairType,
                        BatchNo = line.BatchNo ?? "",
                        CountedBy = line.CountedBy ?? ""
                    }, trans);

                    // Return Target Dates Tracking
                    string insertReturnDtSql = @"INSERT INTO VendIssdDetail_ReturnDTs (VIS_RefID, IssQty, RcvQty, ReturnDT) VALUES (@VIS_RefID, @IssQty, 0, @ReturnDT)";
                    await db.ExecuteAsync(insertReturnDtSql, new { VIS_RefID = lineEntryId, IssQty = line.IssQty, ReturnDT = line.ReturnDT }, trans);

                    // Update VendRcvdDetail IssQty for IAR (Issue After Receive) tracking
                    if (line.IssueRefID > 0)
                    {
                        string updateVrdSql = @"UPDATE VendRcvdDetail SET IssQty = ISNULL(IssQty, 0) + @IssQty WHERE EntryID = @IssueRefID";
                        await db.ExecuteAsync(updateVrdSql, new { IssQty = line.IssQty, IssueRefID = line.IssueRefID }, trans);
                    }

                    // Insert or update VendAssItems
                    string checkAssSql = @"SELECT COUNT(*) FROM VendAssItems WHERE VendID = @VendID AND ProcessID = @ProcessID AND ItemID = @ItemID";
                    int assCount = await db.ExecuteScalarAsync<int>(checkAssSql, new { VendID = header.VendID, ProcessID = header.ProcessID, ItemID = line.ItemCode }, trans);
                    if (assCount == 0)
                    {
                        string insertAssSql = @"INSERT INTO VendAssItems (VendID, ProcessID, ItemID, Rate, Unit, Remarks) VALUES (@VendID, @ProcessID, @ItemID, @Rate, 'Pcs', 'Auto-LotIssuance')";
                        await db.ExecuteAsync(insertAssSql, new { VendID = header.VendID, ProcessID = header.ProcessID, ItemID = line.ItemCode, Rate = line.Rate }, trans);
                    }

                    // Insert into PrintSession
                    string insertPrintSql = @"INSERT INTO PrintSession (RecieptNo) VALUES (@RecieptNo)";
                    await db.ExecuteAsync(insertPrintSql, new { RecieptNo = receiptId }, trans);
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
    }
}
