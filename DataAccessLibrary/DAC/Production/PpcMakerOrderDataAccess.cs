using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Production
{
    public class PpcMakerOrderDataAccess : IPpcMakerOrderDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<PpcMakerOrderDataAccess> _logger;

        public PpcMakerOrderDataAccess(IConfiguration config, ILogger<PpcMakerOrderDataAccess> logger)
        {
            _connectionString = config.GetConnectionString("DefaultConnection")
                ?? config.GetConnectionString("SMBI_AWM")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task EnsureSchemaAsync()
        {
            try
            {
                using var conn = CreateConnection();
                const string sql = @"
                IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'VendIssued_RefID')
                BEGIN
                    ALTER TABLE PPC_Order_Item_Purchases ADD VendIssued_RefID BIGINT NULL;
                END
                IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'MasterPONo')
                BEGIN
                    ALTER TABLE PPC_Order_Item_Purchases ADD MasterPONo VARCHAR(50) NULL;
                END
                IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'IssuedAt')
                BEGIN
                    ALTER TABLE PPC_Order_Item_Purchases ADD IssuedAt DATETIME2 NULL;
                END
                IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'IssuedBy')
                BEGIN
                    ALTER TABLE PPC_Order_Item_Purchases ADD IssuedBy NVARCHAR(100) NULL;
                END";
                await conn.ExecuteAsync(sql);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "EnsureSchemaAsync warning in PpcMakerOrderDataAccess");
            }
        }

        public async Task<PpcMakerPoOrderHeaderDto?> GetOrderHeaderAsync(string orderNo)
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT 
                    co.OrderNo,
                    ISNULL(fc.Name, '') AS CustomerName,
                    ISNULL(co.CustCode, '') AS CustCode,
                    co.DT AS OrderDate,
                    co.DeliveryDT AS DeliveryDate,
                    ISNULL(co.InternalRefNo, '') AS CustomerPoNo,
                    ISNULL(SUM(pip.PurchaseQty), 0) AS TotalPlannedPurchaseQty,
                    COUNT(pip.EntryID) AS TotalPurchaseLinesCount,
                    COUNT(CASE WHEN pip.VendIssued_RefID IS NULL OR pip.MasterPONo IS NULL OR pip.MasterPONo = '' THEN 1 END) AS PendingPurchaseLinesCount
                FROM FCustomerOrders co
                LEFT JOIN ForeignCustomers fc ON co.CustCode = fc.CustCode AND co.Country = fc.Country
                LEFT JOIN PPC_Order_Item_Purchases pip ON co.OrderNo = pip.OrderNo
                WHERE co.OrderNo = @OrderNo
                GROUP BY co.OrderNo, fc.Name, co.CustCode, co.DT, co.DeliveryDT, co.InternalRefNo";

            return await conn.QueryFirstOrDefaultAsync<PpcMakerPoOrderHeaderDto>(sql, new { OrderNo = orderNo });
        }

        public async Task<List<PpcMakerPoOrderHeaderDto>> GetActiveOrdersWithPendingPurchasesAsync()
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT 
                    co.OrderNo,
                    ISNULL(fc.Name, '') AS CustomerName,
                    ISNULL(co.CustCode, '') AS CustCode,
                    co.DT AS OrderDate,
                    co.DeliveryDT AS DeliveryDate,
                    ISNULL(co.InternalRefNo, '') AS CustomerPoNo,
                    ISNULL(SUM(pip.PurchaseQty), 0) AS TotalPlannedPurchaseQty,
                    COUNT(pip.EntryID) AS TotalPurchaseLinesCount,
                    COUNT(CASE WHEN pip.VendIssued_RefID IS NULL OR pip.MasterPONo IS NULL OR pip.MasterPONo = '' THEN 1 END) AS PendingPurchaseLinesCount
                FROM PPC_Order_Item_Purchases pip
                INNER JOIN FCustomerOrders co ON pip.OrderNo = co.OrderNo
                LEFT JOIN ForeignCustomers fc ON co.CustCode = fc.CustCode AND co.Country = fc.Country
                GROUP BY co.OrderNo, fc.Name, co.CustCode, co.DT, co.DeliveryDT, co.InternalRefNo
                ORDER BY PendingPurchaseLinesCount DESC, co.DT DESC";

            var rows = await conn.QueryAsync<PpcMakerPoOrderHeaderDto>(sql);
            return rows.ToList();
        }

        public async Task<List<PpcMakerPoItemRowDto>> GetPpcPurchasesForOrderAsync(string orderNo)
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT 
                    pip.EntryID,
                    pip.OrderNo,
                    pip.ItemID AS CompItemID,
                    ISNULL(i.ItemName, ISNULL(i.Description, pip.ItemID)) AS ItemName,
                    pip.ProcessID,
                    ISNULL(p.Description, '') AS ProcessName,
                    ISNULL(pip.VendID, 0) AS VendID,
                    ISNULL(m.VenderName, 'Unassigned') AS MakerName,
                    pip.PurchaseQty,
                    ISNULL(pip.PurchaseRate, 0) AS PurchaseRate,
                    pip.VendIssued_RefID,
                    pip.MasterPONo,
                    pip.IssuedAt,
                    pip.IssuedBy
                FROM PPC_Order_Item_Purchases pip
                LEFT JOIN Items i ON pip.ItemID = i.ItemID
                LEFT JOIN Processes p ON pip.ProcessID = p.ProcessID
                LEFT JOIN Makers m ON pip.VendID = m.VendID
                WHERE pip.OrderNo = @OrderNo
                ORDER BY pip.EntryID";

            var list = (await conn.QueryAsync<PpcMakerPoItemRowDto>(sql, new { OrderNo = orderNo })).ToList();

            // Populate fallback rates from VendAssItems if Rate is 0 and Maker is assigned
            foreach (var row in list.Where(r => r.PurchaseRate == 0 && r.VendID > 0))
            {
                const string rateSql = @"
                    SELECT TOP 1 Rate 
                    FROM VendAssItems 
                    WHERE VendID = @VendID AND ProcessID = @ProcessID AND ItemID = @ItemID";
                var defaultRate = await conn.ExecuteScalarAsync<decimal?>(rateSql, new
                {
                    row.VendID,
                    row.ProcessID,
                    ItemID = row.CompItemID
                });

                if (defaultRate.HasValue && defaultRate.Value > 0)
                {
                    row.PurchaseRate = defaultRate.Value;
                }
            }

            return list;
        }

        public async Task<List<MakerPOLookupModel>> GetAllMakersAsync()
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT VendID, VendID1, VenderName, ISNULL(MakerCapacity, 0) AS MakerCapacity, ISNULL(Balance, 0) AS Balance 
                FROM VMakerBalances 
                WHERE ISNULL(Active, 1) = 1 
                ORDER BY VenderName";

            var rows = await conn.QueryAsync<MakerPOLookupModel>(sql);
            return rows.ToList();
        }

        public async Task<GenerateMakerPoResult> GenerateMakerPosAsync(GenerateMakerPoRequest request)
        {
            var result = new GenerateMakerPoResult();
            var pendingLines = request.Lines.Where(l => l.IsSelected && !l.IsPosted).ToList();

            if (!pendingLines.Any())
            {
                result.Success = false;
                result.ErrorMessage = "No pending planned purchase items selected.";
                return result;
            }

            // Verify all selected lines have a valid maker assigned
            var unassigned = pendingLines.FirstOrDefault(l => l.VendID <= 0);
            if (unassigned != null)
            {
                result.Success = false;
                result.ErrorMessage = $"Item '{unassigned.CompItemID}' does not have a Maker assigned. Please assign a maker before generating PO.";
                return result;
            }

            using var conn = CreateConnection();
            conn.Open();
            using var trans = conn.BeginTransaction();

            try
            {
                DateTime dt = DateTime.Today;
                string dtFormatted = dt.ToString("ddMMyy");
                string dtSearchFormatted = dt.ToString("dd MMM yy");
                string datePattern = dt.ToString("MMyy");

                // 1. Fetch current max detail receipt counter
                string detailNoSql = @"SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 10) AS INT)) 
                                       FROM VendIssdDetail 
                                       WHERE RecieptID LIKE 'ISU-%'";
                int lastDetailNo = await conn.ExecuteScalarAsync<int?>(detailNoSql, transaction: trans) ?? 0;

                // 2. Fetch current max header receipt counter
                string headerNoSql = @"SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 12) AS INT)) 
                                       FROM VendIssued 
                                       WHERE CONVERT(varchar, DT, 6) = @DtSearch AND RecieptID LIKE 'M-ISU-%'";
                int lastHeaderNo = await conn.ExecuteScalarAsync<int?>(headerNoSql, new { DtSearch = dtSearchFormatted }, trans) ?? 0;

                // 3. Fetch current max MasterPONo sequence
                string masterPoSql = @"SELECT MAX(CAST(LEFT(MasterPONo, CHARINDEX('/', MasterPONo) - 1) AS INT)) 
                                       FROM VendIssued 
                                       WHERE RIGHT(MasterPONo, 4) = @DatePattern AND MasterPONo IS NOT NULL AND CHARINDEX('/', MasterPONo) > 0";
                int maxMasterPoNum = await conn.ExecuteScalarAsync<int?>(masterPoSql, new { DatePattern = datePattern }, trans) ?? 0;

                // 4. Fetch current max Batch No sequence
                string batchNoSql = @"SELECT MAX(CAST(RIGHT(Batch_No, LEN(Batch_No) - 3) AS INT)) 
                                       FROM VendIssdDetail 
                                       WHERE Batch_No LIKE 'B17%' AND ISNUMERIC(RIGHT(Batch_No, LEN(Batch_No) - 3)) = 1";
                int maxBatchNo = await conn.ExecuteScalarAsync<int?>(batchNoSql, transaction: trans) ?? 0;

                // 5. Group lines by Maker (VendID):
                // Rule: Items for the same maker share the same MasterPONo; items for a different maker get a distinct MasterPONo.
                // Rule: Every row gets a unique RecieptID in VendIssued (M-ISU-ddMMyyN) and VendIssdDetail (ISU-ddMMyyN).
                var makerGroups = pendingLines.GroupBy(x => x.VendID).ToList();

                foreach (var group in makerGroups)
                {
                    maxMasterPoNum++;
                    string currentMasterPoNo = $"{maxMasterPoNum}/{datePattern}";
                    result.GeneratedMasterPoNumbers.Add(currentMasterPoNo);

                    foreach (var line in group)
                    {
                        lastHeaderNo++;
                        string headerReceiptId = $"M-ISU-{dtFormatted}{lastHeaderNo}";

                        lastDetailNo++;
                        string detailReceiptId = $"ISU-{dtFormatted}{lastDetailNo}";

                        // 4a. Insert 1 row in VendIssued per item
                        string insertHeaderSql = @"
                            INSERT INTO VendIssued (
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

                        long headerEntryId = await conn.ExecuteScalarAsync<long>(insertHeaderSql, new
                        {
                            VendID = line.VendID,
                            DT = dt,
                            RecieptID = headerReceiptId,
                            UserID = request.UserId,
                            ProcessID = line.ProcessID,
                            ItemID = line.CompItemID,
                            UserName = request.UserName,
                            MachineName = request.MachineName,
                            SpecialInstructions = line.SpecialInstructions ?? string.Empty,
                            ExcessQtyPercentage = 0,
                            MaximumRcvingsAgainstPO = 1,
                            MasterPONo = currentMasterPoNo,
                            SampleProvided = line.SampleProvided,
                            DrawingProvided = line.DrawingProvided,
                            ForgingProvided = line.ForgingProvided,
                            SteelProvided = line.SteelProvided,
                            Authorized = 1,
                            IssEmpID = request.IssEmpID ?? string.Empty,
                            SteelType_RefID = line.SteelType_RefID ?? request.SteelType_RefID
                        }, trans);

                        // 4b. Insert 1 row in VendIssdDetail per item
                        string insertLineSql = @"
                            INSERT INTO VendIssdDetail (
                                RefID, RecieptID, ItemCode, Rate, IssQty, ReqAuth, OrderNo, RcvProcessID,
                                ReturnDT, Priority, Rcvd_RefID, LotNo, ReWorkLot, Repair_RefID, Batch_No,
                                CountedBy, VID_EmpID
                            ) VALUES (
                                @RefID, @RecieptID, @ItemCode, @Rate, @IssQty, 0, @OrderNo, @RcvProcessID,
                                @ReturnDT, @Priority, 0, '0', 0, 0, @Batch_No,
                                @CountedBy, @VID_EmpID
                            );
                            SELECT SCOPE_IDENTITY();";

                        string finalBatch = line.BatchNo;
                        if (string.IsNullOrWhiteSpace(finalBatch) || finalBatch == "B171")
                        {
                            maxBatchNo++;
                            finalBatch = $"B17{maxBatchNo}";
                        }

                        long lineEntryId = await conn.ExecuteScalarAsync<long>(insertLineSql, new
                        {
                            RefID = headerEntryId,
                            RecieptID = detailReceiptId,
                            ItemCode = line.CompItemID,
                            Rate = line.PurchaseRate,
                            IssQty = line.PurchaseQty,
                            OrderNo = request.OrderNo,
                            RcvProcessID = line.ProcessID,
                            ReturnDT = line.ReturnDT,
                            Priority = 2,
                            Batch_No = finalBatch,
                            CountedBy = request.CountedBy ?? string.Empty,
                            VID_EmpID = request.IssEmpID ?? string.Empty
                        }, trans);

                        // 4c. Insert Return Dates
                        string insertReturnDtSql = @"
                            INSERT INTO VendIssdDetail_ReturnDTs (VIS_RefID, IssQty, RcvQty, ReturnDT)
                            VALUES (@VIS_RefID, @IssQty, 0, @ReturnDT)";
                        await conn.ExecuteAsync(insertReturnDtSql, new
                        {
                            VIS_RefID = lineEntryId,
                            IssQty = line.PurchaseQty,
                            ReturnDT = line.ReturnDT
                        }, trans);

                        // 4d. More Details (Checker/CountedBy)
                        if (!string.IsNullOrWhiteSpace(request.CountedBy))
                        {
                            string insertMoreSql = @"
                                INSERT INTO VendIssdDetail_MoreDetails (VID_RefID, EmpID)
                                VALUES (@VID_RefID, @EmpID)";
                            await conn.ExecuteAsync(insertMoreSql, new { VID_RefID = lineEntryId, EmpID = request.CountedBy }, trans);
                        }

                        // 4e. Auth Detail
                        string insertAuthSql = @"
                            INSERT INTO VendIssAuthDetail (UserID, IssID, DT)
                            VALUES (@UserID, @IssID, GETDATE())";
                        await conn.ExecuteAsync(insertAuthSql, new { UserID = request.UserId, IssID = lineEntryId }, trans);

                        // 4f. Auto assign item to maker if not already assigned
                        string checkAssignedSql = @"SELECT COUNT(1) FROM VendAssItems 
                                                    WHERE VendID = @VendID AND ProcessID = @ProcessID AND ItemID = @ItemID";
                        int countAssigned = await conn.ExecuteScalarAsync<int>(checkAssignedSql, new
                        {
                            line.VendID,
                            line.ProcessID,
                            ItemID = line.CompItemID
                        }, trans);

                        if (countAssigned == 0)
                        {
                            string autoAssignSql = @"
                                INSERT INTO VendAssItems (VendID, ProcessID, ItemID, Rate, Unit, Remarks)
                                VALUES (@VendID, @ProcessID, @ItemID, @Rate, 'Pcs', 'Auto')";
                            await conn.ExecuteAsync(autoAssignSql, new
                            {
                                line.VendID,
                                line.ProcessID,
                                ItemID = line.CompItemID,
                                Rate = line.PurchaseRate > 0 ? line.PurchaseRate : 1m
                            }, trans);
                        }

                        // 4g. Execute Inventory update procedure
                        var spParams = new DynamicParameters();
                        spParams.Add("@ItemCode", line.CompItemID);
                        spParams.Add("@ProcessID", line.ProcessID);
                        spParams.Add("@Qty", line.PurchaseQty);
                        spParams.Add("@OrderNo", request.OrderNo);
                        spParams.Add("@Rcvd_RefID", 0);
                        await conn.ExecuteAsync("SP_UpdateForIARNew", spParams, trans, commandType: CommandType.StoredProcedure);

                        // 4h. Insert PrintSession record for detail
                        string insertPrintSessionSql = @"INSERT INTO PrintSession (RecieptNo) VALUES (@RecieptNo)";
                        await conn.ExecuteAsync(insertPrintSessionSql, new { RecieptNo = detailReceiptId }, trans);

                        // 4i. Update PPC_Order_Item_Purchases tracking
                        string updatePpcSql = @"
                            UPDATE PPC_Order_Item_Purchases
                            SET VendIssued_RefID = @VendIssued_RefID,
                                MasterPONo = @MasterPONo,
                                IssuedAt = GETUTCDATE(),
                                IssuedBy = @IssuedBy,
                                VendID = @VendID,
                                PurchaseRate = @PurchaseRate
                            WHERE EntryID = @PpcEntryID";

                        await conn.ExecuteAsync(updatePpcSql, new
                        {
                            VendIssued_RefID = headerEntryId,
                            MasterPONo = currentMasterPoNo,
                            IssuedBy = request.UserName,
                            VendID = line.VendID,
                            PurchaseRate = line.PurchaseRate,
                            PpcEntryID = line.EntryID
                        }, trans);

                        result.TotalRowsCreated++;
                    }
                }

                trans.Commit();
                result.Success = true;
                _logger.LogInformation("Successfully generated {RowCount} Maker PO rows across {PoCount} Master POs for Order #{OrderNo}",
                    result.TotalRowsCreated, result.GeneratedMasterPoNumbers.Count, request.OrderNo);
                return result;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                _logger.LogError(ex, "Failed to generate Maker POs for Order #{OrderNo}", request.OrderNo);
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
        }
    }
}
