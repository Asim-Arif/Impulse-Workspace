using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Stock
{
    public class PpcStockAdjustmentDataAccess : IPpcStockAdjustmentDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<PpcStockAdjustmentDataAccess> _logger;

        public PpcStockAdjustmentDataAccess(IConfiguration config, ILogger<PpcStockAdjustmentDataAccess> logger)
        {
            _connectionString = config.GetConnectionString("DefaultConnection")
                ?? config.GetConnectionString("SMBI_AWM")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private SqlConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<PpcStockAdjustmentOrderHeaderModel?> GetOrderHeaderAsync(string orderNo)
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT 
                    fo.OrderNo,
                    fo.DT AS OrderDate,
                    fo.CustCode AS CustomerCode,
                    ISNULL(fc.Name, fo.CustCode) AS CustomerName,
                    ISNULL(fo.InternalRefNo, '') AS CustPONo,
                    COUNT(pip.EntryID) AS TotalPlannedItems,
                    ISNULL(SUM(pip.StockQty), 0) AS TotalPlannedStockQty,
                    ISNULL(SUM(pip.StockQty) - ISNULL(SUM(soa.AlreadyAdjusted), 0), 0) AS TotalRemainingStockQty
                FROM FCustomerOrders fo
                LEFT JOIN ForeignCustomers fc ON fo.CustCode = fc.CustCode AND fo.Country = fc.Country
                INNER JOIN PPC_Order_Item_Planning pip ON fo.OrderNo = pip.OrderNo AND pip.StockQty > 0
                OUTER APPLY (
                    SELECT SUM(Qty) AS AlreadyAdjusted 
                    FROM StockOrderAdjustments 
                    WHERE OrderNo = pip.OrderNo AND ItemID = pip.ItemID
                ) soa
                WHERE fo.OrderNo = @OrderNo
                GROUP BY fo.OrderNo, fo.DT, fo.CustCode, fc.Name, fo.InternalRefNo";

            return await conn.QueryFirstOrDefaultAsync<PpcStockAdjustmentOrderHeaderModel>(sql, new { OrderNo = orderNo });
        }

        public async Task<List<PpcStockAdjustmentOrderHeaderModel>> GetActiveOrdersWithStockPlanningAsync()
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT 
                    fo.OrderNo,
                    fo.DT AS OrderDate,
                    fo.CustCode AS CustomerCode,
                    ISNULL(fc.Name, fo.CustCode) AS CustomerName,
                    ISNULL(fo.InternalRefNo, '') AS CustPONo,
                    COUNT(pip.EntryID) AS TotalPlannedItems,
                    ISNULL(SUM(pip.StockQty), 0) AS TotalPlannedStockQty,
                    ISNULL(SUM(pip.StockQty) - ISNULL(SUM(soa.AlreadyAdjusted), 0), 0) AS TotalRemainingStockQty
                FROM FCustomerOrders fo
                LEFT JOIN ForeignCustomers fc ON fo.CustCode = fc.CustCode AND fo.Country = fc.Country
                INNER JOIN PPC_Order_Item_Planning pip ON fo.OrderNo = pip.OrderNo AND pip.StockQty > 0
                OUTER APPLY (
                    SELECT SUM(Qty) AS AlreadyAdjusted 
                    FROM StockOrderAdjustments 
                    WHERE OrderNo = pip.OrderNo AND ItemID = pip.ItemID
                ) soa
                GROUP BY fo.OrderNo, fo.DT, fo.CustCode, fc.Name, fo.InternalRefNo
                ORDER BY fo.DT DESC";

            return (await conn.QueryAsync<PpcStockAdjustmentOrderHeaderModel>(sql)).ToList();
        }

        public async Task<List<PpcStockAdjustmentPlannedItemModel>> GetPlannedStockItemsAsync(string orderNo)
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT 
                    pip.EntryID,
                    pip.OrderNo,
                    pip.ItemID,
                    ISNULL(i.ItemName, ISNULL(i.Description, pip.ItemID)) AS ItemName,
                    pip.OrderQty,
                    pip.StockQty AS PlannedStockQty,
                    ISNULL(soa.AlreadyAdjusted, 0) AS AlreadyAdjustedQty,
                    ISNULL(i.InHand, 0) AS AvailableInHandQty
                FROM PPC_Order_Item_Planning pip
                LEFT JOIN Items i ON pip.ItemID = i.ItemID
                OUTER APPLY (
                    SELECT SUM(Qty) AS AlreadyAdjusted 
                    FROM StockOrderAdjustments 
                    WHERE OrderNo = pip.OrderNo AND ItemID = pip.ItemID
                ) soa
                WHERE pip.OrderNo = @OrderNo AND pip.StockQty > 0
                ORDER BY pip.EntryID";

            var items = (await conn.QueryAsync<PpcStockAdjustmentPlannedItemModel>(sql, new { OrderNo = orderNo })).ToList();
            return items;
        }

        public async Task<List<PpcStockShelfAllocationModel>> GetAvailableShelfLotsAsync(string itemId)
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT 
                    Shelf_RefID,
                    ISNULL(StoreName, '') AS StoreName,
                    ISNULL(RackNo, '') AS RackNo,
                    ISNULL(ShelfNo, '') AS ShelfNo,
                    ISNULL(LotNo, '') AS LotNo,
                    ISNULL(Batch_No, '') AS Batch_No,
                    ISNULL(Mill_Certificate_No, '') AS Mill_Certificate_No,
                    ISNULL(EntryID, 0) AS RCV_ISD_RefID,
                    ISNULL(NetQty, 0) AS NetAvailableQty,
                    0 AS AllocatedQty
                FROM VItemsWithLotWiseStock 
                WHERE ItemID = @ItemId AND NetQty > 0 AND ISNULL(Batch_No, '') <> '' 
                ORDER BY Batch_No";

            return (await conn.QueryAsync<PpcStockShelfAllocationModel>(sql, new { ItemId = itemId })).ToList();
        }

        public async Task<List<ItemProcessLookupModel>> GetProcessesForItemAsync(string itemId)
        {
            using var conn = CreateConnection();
            const string sql = @"
                SELECT ProcessID, Description, ItemSNo 
                FROM VItemProcesses 
                WHERE IsExist = @ItemId 
                ORDER BY ItemSNo";

            return (await conn.QueryAsync<ItemProcessLookupModel>(sql, new { ItemId = itemId })).ToList();
        }

        public async Task<PpcStockAdjustmentResultModel> SaveStockAdjustmentAsync(PpcStockAdjustmentSubmissionModel submission)
        {
            var result = new PpcStockAdjustmentResultModel();
            if (submission == null || submission.Items == null || !submission.Items.Any(i => i.AdjQty > 0))
            {
                result.Success = false;
                result.ErrorMessage = "No items specified for stock adjustment.";
                return result;
            }

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var trans = conn.BeginTransaction();

            try
            {
                // 1. Fetch FactoryMaker ID (default 79)
                int factoryMakerId = 79;
                try
                {
                    var valStr = await conn.ExecuteScalarAsync<string>(
                        "SELECT DataValue FROM GeneralData WHERE DataName = 'FactoryMaker'",
                        transaction: trans);
                    if (int.TryParse(valStr, out int parsedId) && parsedId > 0)
                    {
                        factoryMakerId = parsedId;
                    }
                }
                catch
                {
                    factoryMakerId = 79;
                }

                int rcvEntryId = 0;
                string issNo = string.Empty;
                DateTime dt = submission.AdjustmentDate.Date;
                string machineName = string.IsNullOrWhiteSpace(submission.MachineName) ? Environment.MachineName : submission.MachineName;
                string userName = string.IsNullOrWhiteSpace(submission.UserName) ? "system" : submission.UserName;

                foreach (var item in submission.Items.Where(i => i.AdjQty > 0))
                {
                    // Strict Validation: Cannot exceed RemainingQtyToIssue
                    if (item.AdjQty > item.RemainingQtyToIssue)
                    {
                        throw new InvalidOperationException($"Cannot issue {item.AdjQty} for item {item.ItemID}. Maximum remaining planned qty is {item.RemainingQtyToIssue}.");
                    }

                    int totalAlloc = item.Allocations?.Sum(a => a.AllocatedQty) ?? 0;
                    if (totalAlloc != item.AdjQty)
                    {
                        throw new InvalidOperationException($"Total shelf lot allocation ({totalAlloc}) does not match adjustment qty ({item.AdjQty}) for item {item.ItemID}.");
                    }

                    if (!item.SelectedProcessId.HasValue || item.SelectedProcessId.Value <= 0)
                    {
                        throw new InvalidOperationException($"Please select a valid process for item {item.ItemID}.");
                    }

                    int originalProcessId = item.SelectedProcessId.Value;

                    // 2. Call SP_GetPrevProcID
                    var prevProcParams = new DynamicParameters();
                    prevProcParams.Add("@ItemCode", item.ItemID);
                    prevProcParams.Add("@ProcID", originalProcessId);
                    prevProcParams.Add("@PrevProcID", dbType: DbType.Int32, direction: ParameterDirection.Output);
                    await conn.ExecuteAsync("SP_GetPrevProcID", prevProcParams, transaction: trans, commandType: CommandType.StoredProcedure);

                    int prevProcId = prevProcParams.Get<int?>("@PrevProcID") ?? 0;
                    int processId = 0;
                    int nextProcessId = 0;

                    if (prevProcId == 0)
                    {
                        processId = 0;
                        nextProcessId = originalProcessId;
                    }
                    else
                    {
                        processId = prevProcId;
                        // Call SP_GetNextProcID
                        var nextProcParams = new DynamicParameters();
                        nextProcParams.Add("@ItemCode", item.ItemID);
                        nextProcParams.Add("@ProcID", processId);
                        nextProcParams.Add("@ReWorkLot", false);
                        nextProcParams.Add("@RepairType", 0);
                        nextProcParams.Add("@NextProcID", dbType: DbType.Int32, direction: ParameterDirection.Output);
                        await conn.ExecuteAsync("SP_GetNextProcID", nextProcParams, transaction: trans, commandType: CommandType.StoredProcedure);
                        nextProcessId = nextProcParams.Get<int?>("@NextProcID") ?? 0;
                    }

                    // 3. Create VendReceived Header if not already created
                    if (rcvEntryId == 0)
                    {
                        int processToSave = (processId == 0 ? originalProcessId : processId);
                        const string insertRcvSql = @"
                            INSERT INTO VendReceived(VendID, DT, RecieptID, UserID, ProcessID, Issuance_RefID)
                            VALUES(@VendID, @DT, '', 0, @ProcessID, 0);
                            SELECT CAST(SCOPE_IDENTITY() AS INT);";

                        rcvEntryId = await conn.ExecuteScalarAsync<int>(insertRcvSql, new
                        {
                            VendID = factoryMakerId,
                            DT = dt,
                            ProcessID = processToSave
                        }, transaction: trans);
                    }

                    // 4. Group allocations by Batch_No
                    var batchGroups = item.Allocations
                        .Where(a => a.AllocatedQty > 0)
                        .GroupBy(a => a.Batch_No)
                        .ToList();

                    foreach (var bg in batchGroups)
                    {
                        int batchQty = bg.Sum(a => a.AllocatedQty);
                        string batchNo = bg.Key;
                        string milNo = bg.First().Mill_Certificate_No ?? string.Empty;
                        string sourceLotNo = bg.First().LotNo ?? string.Empty;

                        string generatedLotNo;
                        if (batchGroups.Count > 1)
                        {
                            generatedLotNo = await GetNextLotNoAsync(conn, trans, dt);
                        }
                        else
                        {
                            generatedLotNo = await GetNextSubLotNoAsync(conn, trans, sourceLotNo);
                        }

                        // Insert Lots_List
                        const string insertLotSql = @"
                            INSERT INTO Lots_List(LotNo, ItemID, Lot_Type, Reference_LotNo, Batch_No, Mill_Certificate_No)
                            VALUES(@LotNo, @ItemID, 4, '', @Batch_No, @Mill_Certificate_No);";

                        await conn.ExecuteAsync(insertLotSql, new
                        {
                            LotNo = generatedLotNo,
                            ItemID = item.ItemID,
                            Batch_No = batchNo,
                            Mill_Certificate_No = milNo
                        }, transaction: trans);

                        result.GeneratedLots.Add($"Item: {item.ItemID} | Lot #: {generatedLotNo} | Batch: {batchNo} | Qty: {batchQty}");

                        // Insert VendRcvdDetail
                        const string insertVrdSql = @"
                            INSERT INTO VendRcvdDetail(
                                RefID, ItemCode, RecieptID, RcvdQty, Rate, LotNo, NextProcessID, OrderNo,
                                CountedBy, Issue_RefID, ProcessID, RcvdWeight, Opening_RefID
                            ) VALUES (
                                @RefID, @ItemCode, '', @RcvdQty, 0, @LotNo, @NextProcessID, @OrderNo,
                                '', 0, @ProcessID, 0, 0
                            );
                            SELECT CAST(SCOPE_IDENTITY() AS INT);";

                        int vidEntryId = await conn.ExecuteScalarAsync<int>(insertVrdSql, new
                        {
                            RefID = rcvEntryId,
                            ItemCode = item.ItemID,
                            RcvdQty = batchQty,
                            LotNo = generatedLotNo,
                            NextProcessID = nextProcessId,
                            OrderNo = submission.OrderNo,
                            ProcessID = processId
                        }, transaction: trans);

                        // 5. For each allocation line in this batch
                        foreach (var alloc in bg)
                        {
                            // Insert StockOrderAdjustments
                            const string insertSoaSql = @"
                                INSERT INTO StockOrderAdjustments(OrderNo, ItemID, Qty, UserName, MachineName, DTEntry, VID_RefID, Shelf_RefID)
                                VALUES(@OrderNo, @ItemID, @Qty, @UserName, @MachineName, GETDATE(), @VID_RefID, @Shelf_RefID);
                                SELECT CAST(SCOPE_IDENTITY() AS INT);";

                            int soaEntryId = await conn.ExecuteScalarAsync<int>(insertSoaSql, new
                            {
                                OrderNo = submission.OrderNo,
                                ItemID = item.ItemID,
                                Qty = alloc.AllocatedQty,
                                UserName = userName,
                                MachineName = machineName,
                                VID_RefID = vidEntryId,
                                Shelf_RefID = alloc.Shelf_RefID
                            }, transaction: trans);

                            // Insert StockOrderAdjustmentsDetail
                            const string insertSoadSql = @"
                                INSERT INTO StockOrderAdjustmentsDetail(SOA_RefID, Shelf_RefID, Qty, RCV_ISD_RefID)
                                VALUES(@SOA_RefID, @Shelf_RefID, @Qty, @RCV_ISD_RefID);";

                            await conn.ExecuteAsync(insertSoadSql, new
                            {
                                SOA_RefID = soaEntryId,
                                Shelf_RefID = alloc.Shelf_RefID,
                                Qty = alloc.AllocatedQty,
                                RCV_ISD_RefID = alloc.RCV_ISD_RefID
                            }, transaction: trans);

                            // Calculate balance
                            double currentInHand = await conn.ExecuteScalarAsync<double?>(
                                "SELECT InHand FROM Items WHERE ItemID = @ItemID",
                                new { ItemID = item.ItemID }, transaction: trans) ?? 0;

                            double newBalance = currentInHand - alloc.AllocatedQty;
                            string desc = $"Stock Adjusted against Order # {submission.OrderNo}";

                            // Insert ItemLedgerSimple
                            const string insertIlsSql = @"
                                INSERT INTO ItemLedgerSimple(ItemID, DT, Description, Qty, Balance, OrderNo, Location)
                                VALUES(@ItemID, @DT, @Description, @Qty, @Balance, @OrderNo, @Location);
                                SELECT CAST(SCOPE_IDENTITY() AS INT);";

                            int ledgerRefId = await conn.ExecuteScalarAsync<int>(insertIlsSql, new
                            {
                                ItemID = item.ItemID,
                                DT = dt,
                                Description = desc,
                                Qty = -alloc.AllocatedQty,
                                Balance = newBalance,
                                OrderNo = submission.OrderNo,
                                Location = alloc.Shelf_RefID.ToString()
                            }, transaction: trans);

                            // Adjust Item InHand
                            await conn.ExecuteAsync(
                                "UPDATE Items SET InHand = InHand - @Qty WHERE ItemID = @ItemID",
                                new { Qty = alloc.AllocatedQty, ItemID = item.ItemID }, transaction: trans);

                            // Adjust future ItemLedgerSimple records if any
                            await conn.ExecuteAsync(
                                "UPDATE ItemLedgerSimple SET Balance = Balance - @Qty WHERE ItemID = @ItemID AND DT > @DT",
                                new { Qty = alloc.AllocatedQty, ItemID = item.ItemID, DT = dt }, transaction: trans);

                            // Update shelf location stock via SP
                            var updateLocationParams = new DynamicParameters();
                            updateLocationParams.Add("@MaterialID", item.ItemID);
                            updateLocationParams.Add("@LocationID", alloc.Shelf_RefID);
                            updateLocationParams.Add("@Qty", (float)alloc.AllocatedQty);
                            updateLocationParams.Add("@Sample_Type", 0);
                            updateLocationParams.Add("@CustCode", string.Empty);

                            await conn.ExecuteAsync("UpdateItemLocation_SP", updateLocationParams, transaction: trans, commandType: CommandType.StoredProcedure);

                            // IssItemsSimple Header
                            if (string.IsNullOrEmpty(issNo))
                            {
                                issNo = await GetNextIssuanceNoAsync(conn, trans, dt);
                                const string insertIssSql = @"
                                    INSERT INTO IssItemsSimple(IssNo, DT, OrderNo, Department, AttnPerson, UserName, ComputerName, SampleIssuance)
                                    VALUES(@IssNo, @DT, @OrderNo, 'Production', 'Factory', @UserName, @ComputerName, 0);";

                                await conn.ExecuteAsync(insertIssSql, new
                                {
                                    IssNo = issNo,
                                    DT = dt.Add(DateTime.Now.TimeOfDay),
                                    OrderNo = submission.OrderNo,
                                    UserName = userName,
                                    ComputerName = machineName
                                }, transaction: trans);
                            }

                            // Insert IssItemsSimpleDetail
                            const string insertIisdSql = @"
                                INSERT INTO IssItemsSimpleDetail(IssNo, ItemID, Qty, InHand, LedgerRefID, OrderNo, Location, Remarks, RCV_ISD_RefID, LotNo_Manual)
                                VALUES(@IssNo, @ItemID, @Qty, @InHand, @LedgerRefID, @OrderNo, @Location, '', @RCV_ISD_RefID, @LotNo_Manual);
                                SELECT CAST(SCOPE_IDENTITY() AS INT);";

                            int iisdEntryId = await conn.ExecuteScalarAsync<int>(insertIisdSql, new
                            {
                                IssNo = issNo,
                                ItemID = item.ItemID,
                                Qty = alloc.AllocatedQty,
                                InHand = newBalance,
                                LedgerRefID = ledgerRefId,
                                OrderNo = submission.OrderNo,
                                Location = alloc.Shelf_RefID.ToString(),
                                RCV_ISD_RefID = alloc.RCV_ISD_RefID,
                                LotNo_Manual = generatedLotNo
                            }, transaction: trans);

                            // Insert IssItemsSimpleDetail_More
                            const string insertIisdmSql = @"
                                INSERT INTO IssItemsSimpleDetail_More(IISD_RefID, IssdQty, Shelf_RefID, Remarks)
                                VALUES(@IISD_RefID, @IssdQty, @Shelf_RefID, '');";

                            await conn.ExecuteAsync(insertIisdmSql, new
                            {
                                IISD_RefID = iisdEntryId,
                                IssdQty = alloc.AllocatedQty,
                                Shelf_RefID = alloc.Shelf_RefID
                            }, transaction: trans);
                        }
                    }
                }

                trans.Commit();
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                _logger.LogError(ex, "Error executing SaveStockAdjustmentAsync for Order #{OrderNo}", submission.OrderNo);
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        private async Task<string> GetNextSubLotNoAsync(SqlConnection conn, SqlTransaction trans, string baseLotNo)
        {
            string prefix = baseLotNo + "-";
            int prefixLen = prefix.Length;

            const string sql = @"
                SELECT 
                    ISNULL((SELECT MAX(CAST(REPLACE(LotNo, @Prefix, '') AS INT)) 
                            FROM Lots_List 
                            WHERE LEFT(LotNo, @Len) = @Prefix 
                              AND CHARINDEX('-', REPLACE(LotNo, @Prefix, '')) = 0 
                              AND ISNUMERIC(REPLACE(LotNo, @Prefix, '')) = 1), 0) AS LL,
                    ISNULL((SELECT MAX(CAST(REPLACE(LotNo_Manual, @Prefix, '') AS INT)) 
                            FROM StockOrderOpening 
                            WHERE LEFT(LotNo_Manual, @Len) = @Prefix 
                              AND CHARINDEX('-', REPLACE(LotNo_Manual, @Prefix, '')) = 0 
                              AND ISNUMERIC(REPLACE(LotNo_Manual, @Prefix, '')) = 1), 0) AS SOO,
                    ISNULL((SELECT MAX(CAST(REPLACE(LotNo_Manual, @Prefix, '') AS INT)) 
                            FROM RcvItemsSimpleDetail 
                            WHERE LEFT(LotNo_Manual, @Len) = @Prefix 
                              AND CHARINDEX('-', REPLACE(LotNo_Manual, @Prefix, '')) = 0 
                              AND ISNUMERIC(REPLACE(LotNo_Manual, @Prefix, '')) = 1), 0) AS RCV,
                    ISNULL((SELECT MAX(CAST(REPLACE(LotNo, @Prefix, '') AS INT)) 
                            FROM VendRcvdDetail 
                            WHERE LEFT(LotNo, @Len) = @Prefix 
                              AND CHARINDEX('-', REPLACE(LotNo, @Prefix, '')) = 0 
                              AND ISNUMERIC(REPLACE(LotNo, @Prefix, '')) = 1), 0) AS VRD,
                    ISNULL((SELECT MAX(CAST(REPLACE(LotNo, @Prefix, '') AS INT)) 
                            FROM VendIssdDetail 
                            WHERE LEFT(LotNo, @Len) = @Prefix 
                              AND CHARINDEX('-', REPLACE(LotNo, @Prefix, '')) = 0 
                              AND ISNUMERIC(REPLACE(LotNo, @Prefix, '')) = 1), 0) AS VID;";

            var res = await conn.QueryFirstOrDefaultAsync<dynamic>(sql, new { Prefix = prefix, Len = prefixLen }, transaction: trans);
            int maxLot = 0;
            if (res != null)
            {
                int ll = (int)(res.LL ?? 0);
                int soo = (int)(res.SOO ?? 0);
                int rcv = (int)(res.RCV ?? 0);
                int vrd = (int)(res.VRD ?? 0);
                int vid = (int)(res.VID ?? 0);

                maxLot = Math.Max(ll, Math.Max(soo, Math.Max(rcv, Math.Max(vrd, vid))));
            }

            return prefix + (maxLot + 1).ToString();
        }

        private async Task<string> GetNextLotNoAsync(SqlConnection conn, SqlTransaction trans, DateTime dt)
        {
            string year2 = dt.ToString("yy");

            const string sql = @"
                SELECT 
                    ISNULL((SELECT MAX(CAST(SUBSTRING(LotNo, 3, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT)) 
                            FROM Lots_List 
                            WHERE LEFT(LotNo, 2) = @Year2 AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1), 0) AS LL,
                    ISNULL((SELECT MAX(CAST(SUBSTRING(LotNo, 7, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT)) 
                            FROM VendRcvdDetail 
                            WHERE LEFT(LotNo, 2) = @Year2 AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1), 0) AS VRD,
                    ISNULL((SELECT MAX(CAST(SUBSTRING(LotNo, 7, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT)) 
                            FROM VendIssdDetail 
                            WHERE LEFT(LotNo, 2) = @Year2 AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1), 0) AS VID,
                    ISNULL((SELECT MAX(CAST(SUBSTRING(LotNo_Manual, 7, CASE WHEN CHARINDEX('-', LotNo_Manual) > 0 THEN CHARINDEX('-', LotNo_Manual) - 3 ELSE LEN(LotNo_Manual) END) AS INT)) 
                            FROM StockOrderOpening 
                            WHERE LEFT(LotNo_Manual, 2) = @Year2 AND LEN(LotNo_Manual) = 7 AND ISNUMERIC(LotNo_Manual) = 1), 0) AS SOO,
                    ISNULL((SELECT MAX(CAST(SUBSTRING(LotNo_Manual, 7, CASE WHEN CHARINDEX('-', LotNo_Manual) > 0 THEN CHARINDEX('-', LotNo_Manual) - 3 ELSE LEN(LotNo_Manual) END) AS INT)) 
                            FROM RcvItemsSimpleDetail 
                            WHERE LEFT(LotNo_Manual, 2) = @Year2 AND LEN(LotNo_Manual) = 7 AND ISNUMERIC(LotNo_Manual) = 1), 0) AS RCV;";

            var res = await conn.QueryFirstOrDefaultAsync<dynamic>(sql, new { Year2 = year2 }, transaction: trans);
            int maxLot = 0;
            if (res != null)
            {
                int ll = (int)(res.LL ?? 0);
                int vrd = (int)(res.VRD ?? 0);
                int vid = (int)(res.VID ?? 0);
                int soo = (int)(res.SOO ?? 0);
                int rcv = (int)(res.RCV ?? 0);

                maxLot = Math.Max(ll, Math.Max(vrd, Math.Max(vid, Math.Max(soo, rcv))));
            }

            if (maxLot == 0)
            {
                maxLot = 53000;
            }

            return year2 + (maxLot + 1).ToString("D5");
        }

        private async Task<string> GetNextIssuanceNoAsync(SqlConnection conn, SqlTransaction trans, DateTime dt)
        {
            string dateStr = dt.ToString("MM/dd/yy");
            string prefix = "ISS-" + dt.ToString("yyMMdd-");

            const string sql = @"
                SELECT MAX(CAST(RIGHT(IssNo, 3) AS INT)) 
                FROM IssItemsSimple 
                WHERE CONVERT(CHAR(8), DT, 1) = @DateStr";

            var maxVal = await conn.ExecuteScalarAsync<int?>(sql, new { DateStr = dateStr }, transaction: trans);
            int nextSeq = (maxVal ?? 0) + 1;
            return prefix + nextSeq.ToString("D3");
        }
    }
}
