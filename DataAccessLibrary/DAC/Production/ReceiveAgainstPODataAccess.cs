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
    public class ReceiveAgainstPODataAccess : IReceiveAgainstPODataAccess
    {
        private readonly IConfiguration _config;
        private readonly IProductionSequenceDataAccess _sequenceDataAccess;

        public ReceiveAgainstPODataAccess(IConfiguration config, IProductionSequenceDataAccess sequenceDataAccess)
        {
            _config = config;
            _sequenceDataAccess = sequenceDataAccess;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection configuration string is missing.");

        public async Task<MasterPOSearchResultModel?> SearchMasterPOAsync(string receiptId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string searchVal = receiptId.Trim();

            string sql = @"SELECT TOP 1 VI.EntryID, VI.RecieptID, VI.VendID, ISNULL(M.VenderName, 'N/A') AS VendorName, ISNULL(M.VendID1, '') AS VendID1,
                                  VI.ProcessID, ISNULL(P.Description, 'N/A') AS ProcessName, ISNULL(VI.ItemID, '') AS ItemID, VI.DT,
                                  VI.Authorized, ISNULL(VI.Closed, 0) AS Closed, ISNULL(VI.IssEmpID, '') AS IssEmpID,
                                  CASE WHEN M.VenderName LIKE '%FACTORY%' OR M.VendID1 LIKE '%FAC%' OR M.VendID = 79 OR M.VendID = 129 OR M.VendID = (SELECT TOP 1 CAST(DataValue AS BIGINT) FROM GeneralData WHERE DataName = 'FactoryMaker') THEN 1 ELSE 0 END AS IsFactoryMaker
                           FROM VendIssued VI
                           LEFT JOIN VMakers M ON VI.VendID = M.VendID
                           LEFT JOIN Processes P ON VI.ProcessID = P.ProcessID
                           WHERE VI.RecieptID = @SearchVal OR VI.RecieptID = 'M-ISU-' + @SearchVal
                           ORDER BY VI.EntryID DESC";

            return await db.QueryFirstOrDefaultAsync<MasterPOSearchResultModel>(sql, new { SearchVal = searchVal });
        }

        public async Task<List<MasterPOLineItemModel>> GetMasterPOLinesAsync(long vendIssuedEntryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VID.EntryID AS VendIssdDetailEntryID, ISNULL(VID.OrderNo, '') AS OrderNo, VID.ItemCode, ISNULL(I.ItemName, '') AS ItemName,
                                  (SELECT TOP 1 ItemPic FROM Items WHERE ItemID = VID.ItemCode) AS ItemPic,
                                  VID.IssQty, ISNULL(SUM(VRD.RcvdQty), 0) AS PrevRcvdQty, VID.Rate, ISNULL(VID.LotNo, '') AS LotNo,
                                  ISNULL(VID.ReWorkLot, 0) AS ReWorkLot, ISNULL(VID.Repair_RefID, 0) AS RepairType, VID.RcvProcessID AS ReturnProcessID, VID.Priority
                           FROM VendIssdDetail VID
                           LEFT JOIN Items I ON VID.ItemCode = I.ItemID
                           LEFT JOIN VendRcvdDetail VRD ON VID.EntryID = VRD.Issue_RefID
                           WHERE VID.RefID = @VendIssuedEntryId
                           GROUP BY VID.EntryID, VID.OrderNo, VID.ItemCode, I.ItemName, VID.IssQty, VID.Rate, VID.LotNo, VID.ReWorkLot, VID.Repair_RefID, VID.RcvProcessID, VID.Priority
                           ORDER BY VID.EntryID";

            return (await db.QueryAsync<MasterPOLineItemModel>(sql, new { VendIssuedEntryId = vendIssuedEntryId })).ToList();
        }

        public async Task<long> SavePOReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName)
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

                    // Generate LotNo if not present
                    string lineLotNo = line.LotNo;
                    if (string.IsNullOrWhiteSpace(lineLotNo) || lineLotNo == "0")
                    {
                        lineLotNo = await _sequenceDataAccess.GetNextMainLotNoAsync(header.DT);

                        // Insert into Lots_List
                        string insertLotSql = @"INSERT INTO Lots_List (LotNo, ItemID, Lot_Type, Reference_LotNo, Batch_No, Mill_Certificate_No)
                                                VALUES (@LotNo, @ItemID, 0, '', '', @MillCertNo)";

                        await db.ExecuteAsync(insertLotSql, new { LotNo = lineLotNo, ItemID = line.ItemCode, MillCertNo = header.MillCertNo ?? "" }, trans);
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
    }
}
