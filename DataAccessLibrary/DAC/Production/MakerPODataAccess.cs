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
    public class MakerPODataAccess : IMakerPODataAccess
    {
        private readonly IConfiguration _config;

        public MakerPODataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection configuration string is missing.");

        public async Task<List<ProcessPOLookupModel>> GetPurchaseProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessID, Description, SNO 
                           FROM Processes 
                           WHERE ProcessID IN (SELECT ProcessID FROM Processes_Purchase) 
                           ORDER BY SNO";

            return (await db.QueryAsync<ProcessPOLookupModel>(sql)).ToList();
        }

        public async Task<List<MakerPOLookupModel>> GetMakersForProcessAsync(int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VendID, VendID1, VenderName, ISNULL(MakerCapacity, 0) AS MakerCapacity, ISNULL(Balance, 0) AS Balance 
                           FROM VMakerBalances 
                           WHERE VendID IN (SELECT VendID FROM MakerProcesses WHERE ProcessID = @ProcessId) AND ISNULL(Active, 1) = 1 
                           ORDER BY VenderName";

            return (await db.QueryAsync<MakerPOLookupModel>(sql, new { ProcessId = processId })).ToList();
        }

        public async Task<List<AssignedItemLookupModel>> GetAssignedItemsForMakerAndProcessAsync(long vendId, int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT I.ItemID, I.ItemName, ISNULL(VAI.Rate, 0) AS Rate, ISNULL(VAI.Unit, 'Pcs') AS Unit 
                           FROM Items I 
                           INNER JOIN VendAssItems VAI ON I.ItemID = VAI.ItemID 
                           WHERE VAI.VendID = @VendId AND VAI.ProcessID = @ProcessId 
                           ORDER BY I.ItemName";

            return (await db.QueryAsync<AssignedItemLookupModel>(sql, new { VendId = vendId, ProcessId = processId })).ToList();
        }

        public async Task<List<IssuableOrderItemModel>> GetIssuableOrderItemsAsync(int processId, string itemId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            var parameters = new DynamicParameters();
            parameters.Add("@ProcessID", processId);
            parameters.Add("@ItemID", itemId);

            return (await db.QueryAsync<IssuableOrderItemModel>("IssueItems_SP", parameters, commandType: CommandType.StoredProcedure)).ToList();
        }

        public async Task<List<EmployeeLookupModel>> GetEmployeesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EmpID, Name, ISNULL(Designation, '') AS Designation 
                           FROM Employees 
                           WHERE ISNULL(Active, 1) = 1 
                           ORDER BY Name";

            return (await db.QueryAsync<EmployeeLookupModel>(sql)).ToList();
        }

        public async Task<List<SteelTypeLookupModel>> GetSteelTypesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT SteelID, SteelType FROM SteelTypes ORDER BY SteelType";

            return (await db.QueryAsync<SteelTypeLookupModel>(sql)).ToList();
        }

        public async Task<List<MakerBalanceSummaryModel>> GetMakerStockBalancesAsync(long vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessName, ItemCode, ItemName, SUM(RemQty) AS RemQty 
                           FROM VMakerIssItems 
                           WHERE VendID = @VendId AND ISNULL(RemQty, 0) > 0 
                           GROUP BY ProcessName, ItemCode, ItemName";

            return (await db.QueryAsync<MakerBalanceSummaryModel>(sql, new { VendId = vendId })).ToList();
        }

        public async Task<List<MakerHistorySummaryModel>> GetMakerIssuanceHistoryAsync(long vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT TOP 10 EntryID, ProcessName, ItemCode, ItemName, DT, IssQty, Rate, IssValue, ReturnDT, Priority 
                           FROM VMakerIssItems 
                           WHERE VendID = @VendId 
                           ORDER BY DT DESC, EntryID DESC";

            return (await db.QueryAsync<MakerHistorySummaryModel>(sql, new { VendId = vendId })).ToList();
        }

        public async Task<string> GetNextMasterPONoAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            DateTime today = DateTime.Today;
            string datePattern = today.ToString("MMyy");

            string sql = @"SELECT MAX(CAST(LEFT(MasterPONo, CHARINDEX('/', MasterPONo) - 1) AS INT)) 
                           FROM VendIssued 
                           WHERE RIGHT(MasterPONo, 4) = @DatePattern AND MasterPONo IS NOT NULL AND CHARINDEX('/', MasterPONo) > 0";

            int maxNo = await db.ExecuteScalarAsync<int?>(sql, new { DatePattern = datePattern }) ?? 0;
            return $"{maxNo + 1}/{datePattern}";
        }

        public async Task<string> GetNextBatchNoAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT MAX(CAST(RIGHT(Batch_No, LEN(Batch_No) - 3) AS INT)) 
                           FROM VendIssdDetail 
                           WHERE Batch_No LIKE 'B17%' AND ISNUMERIC(RIGHT(Batch_No, LEN(Batch_No) - 3)) = 1";

            int maxBatch = await db.ExecuteScalarAsync<int?>(sql) ?? 0;
            return $"B17{maxBatch + 1}";
        }

        public async Task<long> SaveMakerPOAsync(CreateMakerPOHeaderModel header, List<CreateMakerPOLineModel> lines, string userName, int userId, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                string masterPoNo = await GetNextMasterPONoAsync();

                // 1. Generate Receipt ID for VendIssued Header
                DateTime dt = header.DT;
                string dtFormatted = dt.ToString("ddMMyy");
                string dtSearchFormatted = dt.ToString("dd MMM yy");

                string headerNoSql = @"SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 12) AS INT)) 
                                       FROM VendIssued 
                                       WHERE CONVERT(varchar, DT, 6) = @DtSearch AND RecieptID LIKE 'M-ISU-%'";

                int lastHeaderNo = await db.ExecuteScalarAsync<int?>(headerNoSql, new { DtSearch = dtSearchFormatted }, trans) ?? 0;
                string headerReceiptId = $"M-ISU-{dtFormatted}{lastHeaderNo + 1}";

                // 2. Insert Header into VendIssued (All 19 Fields)
                string insertHeaderSql = @"INSERT INTO VendIssued (
                                                VendID, DT, RecieptID, UserID, ProcessID, ItemID, UserName, MachineName,
                                                SpecialInstructions, ExcessQtyPercentage, MaximumRcvingsAgainstPO, MasterPONo,
                                                SampleProvided, DrawingProvided, ForgingProvided, SteelProvided, Authorized,
                                                IssEmpID, SteelType_RefID
                                           ) VALUES (
                                                @VendID, @DT, @RecieptID, @UserID, @ProcessID, @ItemID, @UserName, @MachineName,
                                                @SpecialInstructions, @ExcessQtyPercentage, @MaximumRcvingsAgainstPO, @MasterPONo,
                                                @SampleProvided, @DrawingProvided, @ForgingProvided, @SteelProvided, @Authorized,
                                                @IssEmpID, @SteelType_RefID
                                           );
                                           SELECT SCOPE_IDENTITY();";

                long headerEntryId = await db.ExecuteScalarAsync<long>(insertHeaderSql, new
                {
                    header.VendID,
                    header.DT,
                    RecieptID = headerReceiptId,
                    UserID = userId,
                    header.ProcessID,
                    header.ItemID,
                    UserName = userName,
                    MachineName = machineName,
                    header.SpecialInstructions,
                    header.ExcessQtyPercentage,
                    header.MaximumRcvingsAgainstPO,
                    MasterPONo = masterPoNo,
                    header.SampleProvided,
                    header.DrawingProvided,
                    header.ForgingProvided,
                    header.SteelProvided,
                    Authorized = 1,
                    header.IssEmpID,
                    header.SteelType_RefID
                }, trans);

                // 3. Generate Receipt ID base for Detail Lines
                string detailNoSql = @"SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 10) AS INT)) 
                                       FROM VendIssdDetail 
                                       WHERE RecieptID LIKE 'ISU-%'";

                int lastDetailNo = await db.ExecuteScalarAsync<int?>(detailNoSql, transaction: trans) ?? 0;

                // 4. Insert Lines into VendIssdDetail (All 17 Fields)
                foreach (var line in lines)
                {
                    lastDetailNo++;
                    string detailReceiptId = $"ISU-{dtFormatted}{lastDetailNo}";

                    string insertLineSql = @"INSERT INTO VendIssdDetail (
                                                RefID, RecieptID, ItemCode, Rate, IssQty, ReqAuth, OrderNo, RcvProcessID,
                                                ReturnDT, Priority, Rcvd_RefID, LotNo, ReWorkLot, Repair_RefID, Batch_No,
                                                CountedBy, VID_EmpID
                                             ) VALUES (
                                                @RefID, @RecieptID, @ItemCode, @Rate, @IssQty, 0, @OrderNo, @RcvProcessID,
                                                @ReturnDT, @Priority, 0, '0', 0, 0, @Batch_No,
                                                @CountedBy, @VID_EmpID
                                             );
                                             SELECT SCOPE_IDENTITY();";

                    long lineEntryId = await db.ExecuteScalarAsync<long>(insertLineSql, new
                    {
                        RefID = headerEntryId,
                        RecieptID = detailReceiptId,
                        ItemCode = line.ItemCode,
                        Rate = line.Rate,
                        IssQty = line.IssQty,
                        OrderNo = line.OrderNo,
                        RcvProcessID = line.ReturnProcessID,
                        ReturnDT = line.ReturnDT,
                        Priority = line.Priority,
                        Batch_No = line.BatchNo,
                        CountedBy = line.CountedBy,
                        VID_EmpID = line.IssEmpID
                    }, trans);

                    // Insert Return Date record
                    string insertReturnDtSql = @"INSERT INTO VendIssdDetail_ReturnDTs (VIS_RefID, IssQty, RcvQty, ReturnDT)
                                                 VALUES (@VIS_RefID, @IssQty, 0, @ReturnDT)";

                    await db.ExecuteAsync(insertReturnDtSql, new
                    {
                        VIS_RefID = lineEntryId,
                        IssQty = line.IssQty,
                        ReturnDT = line.ReturnDT
                    }, trans);

                    // Insert More Details if Checker Employee exists
                    if (!string.IsNullOrWhiteSpace(header.CheckerEmpID))
                    {
                        string insertMoreSql = @"INSERT INTO VendIssdDetail_MoreDetails (VID_RefID, EmpID)
                                                 VALUES (@VID_RefID, @EmpID)";

                        await db.ExecuteAsync(insertMoreSql, new { VID_RefID = lineEntryId, EmpID = header.CheckerEmpID }, trans);
                    }

                    // Insert Auth Detail
                    string insertAuthSql = @"INSERT INTO VendIssAuthDetail (UserID, IssID, DT)
                                             VALUES (@UserID, @IssID, GETDATE())";

                    await db.ExecuteAsync(insertAuthSql, new { UserID = userId, IssID = lineEntryId }, trans);

                    // Auto assign item to maker if not previously assigned
                    string checkAssignedSql = @"SELECT COUNT(1) FROM VendAssItems 
                                                WHERE VendID = @VendID AND ProcessID = @ProcessID AND ItemID = @ItemID";

                    int countAssigned = await db.ExecuteScalarAsync<int>(checkAssignedSql, new
                    {
                        header.VendID,
                        header.ProcessID,
                        ItemID = line.ItemCode
                    }, trans);

                    if (countAssigned == 0)
                    {
                        string autoAssignSql = @"INSERT INTO VendAssItems (VendID, ProcessID, ItemID, Rate, Unit, Remarks)
                                                 VALUES (@VendID, @ProcessID, @ItemID, @Rate, 'Pcs', 'Auto')";

                        await db.ExecuteAsync(autoAssignSql, new
                        {
                            header.VendID,
                            header.ProcessID,
                            ItemID = line.ItemCode,
                            Rate = line.Rate > 0 ? line.Rate : 1m
                        }, trans);
                    }

                    // Execute inventory update stored procedure
                    var spParams = new DynamicParameters();
                    spParams.Add("@ItemCode", line.ItemCode);
                    spParams.Add("@ProcessID", header.ProcessID);
                    spParams.Add("@Qty", line.IssQty);
                    spParams.Add("@OrderNo", line.OrderNo);
                    spParams.Add("@Rcvd_RefID", 0);

                    await db.ExecuteAsync("SP_UpdateForIARNew", spParams, trans, commandType: CommandType.StoredProcedure);

                    // Insert PrintSession record
                    string insertPrintSessionSql = @"INSERT INTO PrintSession (RecieptNo) VALUES (@RecieptNo)";
                    await db.ExecuteAsync(insertPrintSessionSql, new { RecieptNo = detailReceiptId }, trans);
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
