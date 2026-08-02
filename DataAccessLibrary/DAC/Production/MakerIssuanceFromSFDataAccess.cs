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
    public class MakerIssuanceFromSFDataAccess : IMakerIssuanceFromSFDataAccess
    {
        private readonly IConfiguration _config;
        private readonly IProductionSequenceDataAccess _sequenceDataAccess;

        public MakerIssuanceFromSFDataAccess(IConfiguration config, IProductionSequenceDataAccess sequenceDataAccess)
        {
            _config = config;
            _sequenceDataAccess = sequenceDataAccess;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection configuration string is missing.");

        public async Task<List<ProcessPOLookupModel>> GetAllProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessID, Description, SNO 
                           FROM Processes 
                           ORDER BY SNO";

            return (await db.QueryAsync<ProcessPOLookupModel>(sql)).ToList();
        }

        public async Task<SFLotLookupResultModel?> LookupLotNoAsync(string lotNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT SOO.EntryID, ISNULL(SOO.LotNo_Manual, '') AS LotNo, (SOO.Qty - ISNULL(SOO.IssdQty, 0)) AS RcvdQty, 
                                  SOO.ProcessID, P.Description AS ProcessName, SOO.ItemID AS ItemCode, I.ItemName,
                                  SOO.Shelf_RefID, SOO.EntryID AS SOO_RefID, (SOO.Qty - ISNULL(SOO.IssdQty, 0)) AS SplitQty, ISNULL(SOO.Remarks, '') AS Remarks 
                           FROM VStockOrderOpening SOO
                           LEFT JOIN Processes P ON SOO.ProcessID = P.ProcessID
                           LEFT JOIN Items I ON SOO.ItemID = I.ItemID
                           WHERE SOO.LotNo_Manual = @LotNo AND (SOO.Qty - ISNULL(SOO.IssdQty, 0)) > 0";

            return await db.QueryFirstOrDefaultAsync<SFLotLookupResultModel>(sql, new { LotNo = lotNo });
        }

        public async Task<List<SFStockItemLookupModel>> GetAvailableSFItemsAsync(int processId = 0, long vendId = 0, int returnProcessId = 0)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT DISTINCT I.ItemID, I.ItemName 
                           FROM Items I
                           INNER JOIN VStockOrderOpening SOO ON I.ItemID = SOO.ItemID
                           WHERE (SOO.Qty - ISNULL(SOO.IssdQty, 0)) > 0
                             AND (@ProcessId = 0 OR SOO.ProcessID = @ProcessId)
                             AND (@VendId = 0 OR I.ItemID IN (
                                 SELECT ItemID FROM VendAssItems 
                                 WHERE VendID = @VendId 
                                   AND (@ReturnProcessId = 0 OR ProcessID = @ReturnProcessId)
                             ))
                           ORDER BY I.ItemID";

            return (await db.QueryAsync<SFStockItemLookupModel>(sql, new { ProcessId = processId, VendId = vendId, ReturnProcessId = returnProcessId })).ToList();
        }

        public async Task<List<SFStockOpeningLocationModel>> GetStockOpeningLocationsForItemAsync(string itemId, int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT SOO.EntryID, SOO.ItemID, SOO.ProcessID, SOO.Shelf_RefID,
                                  ISNULL(S.StoreName, 'Main Store') AS StoreName, ISNULL(S.RackNo, 'N/A') AS RackNo, ISNULL(S.ShelfNo, 'N/A') AS ShelfNo,
                                  (SOO.Qty - ISNULL(SOO.IssdQty, 0)) AS InHandQty, ISNULL(SOO.LotNo_Manual, '') AS LotNo, ISNULL(SOO.Remarks, '') AS Remarks
                           FROM VStockOrderOpening SOO
                           LEFT JOIN VStoreShelfs S ON SOO.Shelf_RefID = S.EntryID
                           WHERE SOO.ItemID = @ItemId AND (SOO.Qty - ISNULL(SOO.IssdQty, 0)) > 0
                             AND (@ProcessId = 0 OR SOO.ProcessID = @ProcessId)
                           ORDER BY SOO.EntryID DESC";

            return (await db.QueryAsync<SFStockOpeningLocationModel>(sql, new { ItemId = itemId, ProcessId = processId })).ToList();
        }

        public async Task<List<UnshippedOrderLookupModel>> GetUnshippedOrdersForItemAsync(string itemId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT DISTINCT OrderNo, ISNULL(InternalRefNo, '') AS InternalRefNo
                           FROM VrptOrders_ForProduction
                           WHERE CompItemID = @ItemId AND OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList)
                           ORDER BY OrderNo";

            return (await db.QueryAsync<UnshippedOrderLookupModel>(sql, new { ItemId = itemId })).ToList();
        }

        public async Task<long> SaveSFIssuanceAsync(CreateSFIssuanceHeaderModel header, List<CreateSFIssuanceLineModel> lines, string userName, int userId, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                // 1. Generate Header Receipt ID
                string headerReceiptId = await _sequenceDataAccess.GetNextHeaderReceiptIDAsync(header.DT);

                // 2. Insert Header into VendIssued (All 19 Fields)
                string insertHeaderSql = @"INSERT INTO VendIssued (
                                                VendID, DT, RecieptID, UserID, ProcessID, ItemID, UserName, MachineName,
                                                SpecialInstructions, ExcessQtyPercentage, MaximumRcvingsAgainstPO, MasterPONo,
                                                SampleProvided, DrawingProvided, ForgingProvided, SteelProvided, Authorized,
                                                IssEmpID, SteelType_RefID
                                           ) VALUES (
                                                @VendID, @DT, @RecieptID, @UserID, @ProcessID, @ItemID, @UserName, @MachineName,
                                                @SpecialInstructions, 0, 1, NULL,
                                                0, 0, 0, 0, 1,
                                                @IssEmpID, 0
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
                    header.IssEmpID
                }, trans);

                // 3. Process Lines
                foreach (var line in lines)
                {
                    string detailReceiptId = await _sequenceDataAccess.GetNextDetailReceiptIDAsync(header.DT);

                    // Generate sub lot number
                    string subLotNo = await _sequenceDataAccess.GetNextSubLotNoAsync(line.LotNo);

                    // Fetch Batch & Mill Certificate from original Lot
                    string batchNo = "";
                    string millNo = "";
                    if (!string.IsNullOrWhiteSpace(line.LotNo))
                    {
                        string lotSql = @"SELECT TOP 1 ISNULL(Batch_No, '') AS BatchNo, ISNULL(Mill_Certificate_No, '') AS MillNo 
                                          FROM Lots_List WHERE LotNo = @LotNo";
                        var lotObj = await db.QueryFirstOrDefaultAsync<dynamic>(lotSql, new { line.LotNo }, trans);
                        if (lotObj != null)
                        {
                            batchNo = lotObj.BatchNo ?? "";
                            millNo = lotObj.MillNo ?? "";
                        }
                    }

                    // Insert into Lots_List
                    string insertLotSql = @"INSERT INTO Lots_List (LotNo, ItemID, Lot_Type, Reference_LotNo, Batch_No, Mill_Certificate_No)
                                            VALUES (@LotNo, @ItemID, 3, '', @BatchNo, @MillNo)";

                    await db.ExecuteAsync(insertLotSql, new
                    {
                        LotNo = subLotNo,
                        ItemID = line.ItemCode,
                        BatchNo = batchNo,
                        MillNo = millNo
                    }, trans);

                    // Insert into VendIssdDetail (All 17 Fields)
                    string insertLineSql = @"INSERT INTO VendIssdDetail (
                                                RefID, RecieptID, ItemCode, Rate, IssQty, ReqAuth, OrderNo, RcvProcessID,
                                                ReturnDT, Priority, Rcvd_RefID, LotNo, ReWorkLot, Repair_RefID, Batch_No,
                                                CountedBy, VID_EmpID
                                             ) VALUES (
                                                @RefID, @RecieptID, @ItemCode, @Rate, @IssQty, 0, @OrderNo, @RcvProcessID,
                                                @ReturnDT, @Priority, 0, @LotNo, 0, 0, @BatchNo,
                                                '', @VID_EmpID
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
                        LotNo = subLotNo,
                        BatchNo = batchNo,
                        VID_EmpID = header.IssEmpID
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

                    // Update VendRcvdDetail
                    string updateVrdSql = @"UPDATE VendRcvdDetail 
                                            SET IssQty = ISNULL(IssQty, 0) + @IssQty 
                                            WHERE Opening_RefID = @OpeningRefID";

                    await db.ExecuteAsync(updateVrdSql, new { line.IssQty, OpeningRefID = line.OpeningEntryID }, trans);

                    // Update StockOrderOpening Remarks
                    if (!string.IsNullOrWhiteSpace(line.Remarks))
                    {
                        string updateSooSql = @"UPDATE StockOrderOpening SET Remarks = @Remarks WHERE EntryID = @OpeningRefID";
                        await db.ExecuteAsync(updateSooSql, new { line.Remarks, OpeningRefID = line.OpeningEntryID }, trans);
                    }

                    // Insert into StockOrderOpening_Issuance
                    string insertSooIssSql = @"INSERT INTO StockOrderOpening_Issuance (SOO_RefID, QtyIssd, VID_RefID)
                                               VALUES (@SOO_RefID, @QtyIssd, @VID_RefID)";

                    await db.ExecuteAsync(insertSooIssSql, new
                    {
                        SOO_RefID = line.OpeningEntryID,
                        QtyIssd = line.IssQty,
                        VID_RefID = lineEntryId
                    }, trans);

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
