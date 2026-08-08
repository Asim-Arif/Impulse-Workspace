using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Dashboard;
using DataAccessLibrary.Models.ViewModels.Dashboard;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Dashboard
{
    public class ProductionPlanningDashboardDataAccess : IProductionPlanningDashboardDataAccess
    {
        private readonly IConfiguration _config;

        public ProductionPlanningDashboardDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<List<LookupItemString>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT '0' AS Id, '<Comp>' AS Name
                UNION ALL
                SELECT CustCode AS Id, CustCode AS Name
                FROM ForeignCustomers
                WHERE Active = 1
                ORDER BY Id";
            var result = await db.QueryAsync<LookupItemString>(sql);
            return result.ToList();
        }

        public async Task<List<string>> GetCountriesByCustomerAsync(string custCode)
        {
            if (string.IsNullOrWhiteSpace(custCode) || custCode == "0")
                return new List<string> { "<Comp>" };

            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT DISTINCT ISNULL(Country, '<Comp>') AS Country
                FROM ForeignCustomers
                WHERE CustCode = @CustCode
                ORDER BY Country";
            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode });
            var list = result.ToList();
            if (!list.Any()) list.Add("<Comp>");
            return list;
        }

        public async Task<List<LookupItemString>> GetArticlesAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            if (string.IsNullOrWhiteSpace(custCode) || custCode == "0")
            {
                const string sql = @"
                    SELECT ItemID AS Id, '{' + ItemID + '} ' + ISNULL(ItemName, '') AS Name
                    FROM Items
                    WHERE InActive = 0
                    ORDER BY ItemID";
                var result = await db.QueryAsync<LookupItemString>(sql);
                return result.ToList();
            }
            else
            {
                const string sql = @"
                    SELECT CompItemID AS Id,
                           '{' + CompItemID + '} ' + ISNULL(ItemID, '') + 
                           CASE WHEN ISNULL(OCR, '') <> '' THEN ' {' + OCR + '}' ELSE '' END +
                           CASE WHEN ISNULL(OCR2, '') <> '' THEN ' {' + OCR2 + '}' ELSE '' END +
                           ' ' + ISNULL(CompDesc, '') AS Name
                    FROM VFCustomerCatalog
                    WHERE CustCode = @CustCode AND (@Country = '<Comp>' OR Country = @Country)
                    ORDER BY CompItemID";
                var result = await db.QueryAsync<LookupItemString>(sql, new { CustCode = custCode, Country = country });
                return result.ToList();
            }
        }

        public async Task<ProductionPlanningDashboardData> GetDashboardDataAsync(string itemCode, string? filterOrderNo = null)
        {
            var data = new ProductionPlanningDashboardData { ItemCode = itemCode };
            if (string.IsNullOrWhiteSpace(itemCode)) return data;

            using IDbConnection db = new SqlConnection(ConnectionString);

            // 1. Item Name & Min Stock
            var itemMeta = await db.QueryFirstOrDefaultAsync<(string? ItemName, long? MinLevel)>(
                "SELECT ItemName, MinLevel FROM Items WHERE ItemID = @ItemCode",
                new { ItemCode = itemCode });
            data.ItemName = itemMeta.ItemName ?? itemCode;
            data.Summary.MinStock = itemMeta.MinLevel ?? 0;

            // 2. Customer Orders
            string custOrderSql = @"
                SELECT ID, CustCode, OrderNo, ISNULL(InternalRefNo, OrderNo) AS InternalRefNo,
                       DT, Qty, ShippedQty, DeliveryDT, ISNULL(DeliveryStatus, 0) AS DeliveryStatus,
                       ISNULL(Remarks, '') AS Remarks
                FROM VFOrderItems
                WHERE CompItemCode = @ItemCode
                  AND OrderNo NOT IN ('Stock-Order', 'Stock-OrderIss')
                  AND OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList)
                  AND Qty > ShippedQty";

            if (!string.IsNullOrWhiteSpace(filterOrderNo))
                custOrderSql += " AND OrderNo = @FilterOrderNo";
            custOrderSql += " ORDER BY DeliveryDT ASC, DT ASC";

            var custOrders = await db.QueryAsync<CustOrderDemandItem>(custOrderSql, new { ItemCode = itemCode, FilterOrderNo = filterOrderNo });
            data.CustomerOrders = custOrders.ToList();
            data.Summary.TotalCustOrders = data.CustomerOrders.Sum(x => x.UnshippedQty);

            // 3. Running Lots
            string runningLotsSql = @"
                SELECT T1.LotNo, T1.Description, T1.Qty, T1.OrderNo, ISNULL(T1.InternalRefNo, T1.OrderNo) AS InternalRefNo,
                       FCustomerOrders.CustCode, VLotWithR2InDate.DT AS R2InDT,
                       CASE WHEN ISNULL(TSOOI.VID_RefID, 0) > 0 THEN 'Semi' ELSE '' END AS OriginType,
                       T1.IssuanceEntryID, T1.RcvEntryID, T1.ProcessID
                FROM (
                    SELECT ItemCode, LotNo, OrderNo, Description, InternalRefNo, IssQty - RcvdQty AS Qty, ProcessID,
                           EntryID AS IssuanceEntryID, 0 AS RcvEntryID
                    FROM VVendIssdDetail_ForRunningLots WITH (NOLOCK)
                    WHERE ItemCode = @ItemCode AND LotNo <> '0' AND IssQty - RcvdQty > 0
                      AND EntryID NOT IN (SELECT Issue_RefID FROM VendRcvdDetail WHERE Issue_RefID IS NOT NULL)";

            if (!string.IsNullOrWhiteSpace(filterOrderNo))
                runningLotsSql += " AND OrderNo = @FilterOrderNo";

            runningLotsSql += @"
                    UNION ALL
                    SELECT ItemCode, LotNo, OrderNo, Description, InternalRefNo,
                           RcvdQty - IssQty - ISNULL(Wastage,0) - ISNULL(ReWorkQty,0) AS Qty, ProcessID,
                           0 AS IssuanceEntryID, EntryID AS RcvEntryID
                    FROM VVendRcvdDetail_Simple WITH (NOLOCK)
                    WHERE ItemCode = @ItemCode AND LotNo <> '0'
                      AND RcvdQty - IssQty - ISNULL(Wastage,0) - ISNULL(ReWorkQty,0) > 0
                      AND EntryID NOT IN (SELECT Rcvd_RefID FROM VendIssdDetail WHERE Rcvd_RefID IS NOT NULL)
                      AND ISNULL(Opening_RefID, 0) = 0";

            if (!string.IsNullOrWhiteSpace(filterOrderNo))
                runningLotsSql += " AND OrderNo = @FilterOrderNo";

            runningLotsSql += @") T1
                LEFT OUTER JOIN VLotWithR2InDate ON T1.LotNo = VLotWithR2InDate.LotNo
                LEFT OUTER JOIN VLotWithFirstProcessOfIssuance ON T1.LotNo = VLotWithFirstProcessOfIssuance.LotNo
                LEFT OUTER JOIN (SELECT VID_RefID FROM StockOrderOpening_Issuance GROUP BY VID_RefID) TSOOI ON VLotWithFirstProcessOfIssuance.EntryID = TSOOI.VID_RefID
                LEFT OUTER JOIN FCustomerOrders ON T1.OrderNo = FCustomerOrders.OrderNo
                WHERE T1.LotNo NOT IN (SELECT LotNo FROM Lots_Closed)
                ORDER BY T1.ProcessID";

            var runningLots = await db.QueryAsync<RunningLotItem>(runningLotsSql, new { ItemCode = itemCode, FilterOrderNo = filterOrderNo });
            data.RunningLots = runningLots.ToList();
            data.Summary.TotalRunningLots = data.RunningLots.Sum(x => x.Qty);

            // 4. Open Maker POs
            string openPoSql = @"
                SELECT EntryID, RecieptID, VenderName AS MakerName, OrderNo, Qty, ReturnDT, MasterPONo
                FROM VItemOpenPOs
                WHERE ItemCode = @ItemCode";

            if (!string.IsNullOrWhiteSpace(filterOrderNo))
                openPoSql += " AND OrderNo = @FilterOrderNo";
            openPoSql += " ORDER BY ReturnDT ASC";

            var openPos = await db.QueryAsync<OpenPOItem>(openPoSql, new { ItemCode = itemCode, FilterOrderNo = filterOrderNo });
            data.OpenPOs = openPos.ToList();
            data.Summary.TotalOpenPO = data.OpenPOs.Sum(x => x.Qty);

            // 5. Semi-Finish Stocks & Lookalikes
            const string sfSql = @"
                SELECT Description, Qty - IssdQty AS Qty,
                       ISNULL(StoreName, '') + ' (' + ISNULL(RackNo, '') + ') {' + ISNULL(ShelfNo, '') + '}' AS StoreName,
                       ISNULL(LotNo_Manual, '') AS LotNoManual, ISNULL(Remarks, '') AS Remarks, ItemID, 0 AS IsLookAlike
                FROM VStockOrderOpening
                WHERE ItemID = @ItemCode AND Qty > IssdQty";
            var sfList = await db.QueryAsync<SemiFinishStockItem>(sfSql, new { ItemCode = itemCode });
            data.SemiFinishStocks = sfList.ToList();
            data.Summary.TotalSFStock = data.SemiFinishStocks.Sum(x => x.Qty);

            const string sfLaSql = @"
                SELECT Description, Qty - IssdQty AS Qty,
                       ISNULL(StoreName, '') + ' (' + ISNULL(RackNo, '') + ') {' + ISNULL(ShelfNo, '') + '}' AS StoreName,
                       ISNULL(LotNo_Manual, '') AS LotNoManual, ISNULL(Remarks, '') AS Remarks, ItemID, 1 AS IsLookAlike
                FROM VStockOrderOpening
                WHERE ItemID IN (SELECT ItemID_LookAlike FROM ItemLookalikes WHERE ItemID = @ItemCode) AND Qty > IssdQty";
            var sfLaList = await db.QueryAsync<SemiFinishStockItem>(sfLaSql, new { ItemCode = itemCode });
            data.SemiFinishLookalikes = sfLaList.ToList();

            // 6. Ready-Finish Stocks & Lookalikes
            const string rfSql = @"
                SELECT ItemID, SUM(NetQty) AS NetQty,
                       ISNULL(StoreName, '') + ' (' + ISNULL(RackNo, '') + ') {' + ISNULL(ShelfNo, '') + '}' AS StoreName,
                       ISNULL(Remarks, '') AS Remarks, 0 AS IsLookAlike
                FROM VItemsWithShelfWiseStock
                WHERE ItemID = @ItemCode
                GROUP BY ItemID, StoreName, RackNo, ShelfNo, Remarks
                HAVING SUM(NetQty) > 0";
            var rfList = await db.QueryAsync<ReadyFinishStockItem>(rfSql, new { ItemCode = itemCode });
            data.ReadyFinishStocks = rfList.ToList();
            data.Summary.TotalRFStock = data.ReadyFinishStocks.Sum(x => x.NetQty);

            const string rfLaSql = @"
                SELECT ItemID, SUM(NetQty) AS NetQty,
                       ISNULL(StoreName, '') + ' (' + ISNULL(RackNo, '') + ') {' + ISNULL(ShelfNo, '') + '}' AS StoreName,
                       ISNULL(Remarks, '') AS Remarks, 1 AS IsLookAlike
                FROM VItemsWithShelfWiseStock
                WHERE ItemID IN (SELECT ItemID_LookAlike FROM ItemLookalikes WHERE ItemID = @ItemCode)
                GROUP BY ItemID, StoreName, RackNo, ShelfNo, Remarks
                HAVING SUM(NetQty) > 0";
            var rfLaList = await db.QueryAsync<ReadyFinishStockItem>(rfLaSql, new { ItemCode = itemCode });
            data.ReadyFinishLookalikes = rfLaList.ToList();

            // 7. Forgings / Raw Materials
            const string forgSql = @"
                SELECT VMaterialLocationWiseBalances.RMID, MaterialID, ShelfQty,
                       ISNULL(StoreName, '') + ' - ' + ISNULL(RackNo, '') + ' - ' + ISNULL(ShelfNo, '') AS LocationDisplay,
                       CASE WHEN ISNULL(ItemsRMComp.Functional_Status, 0) = 1 THEN 1 ELSE 0 END AS IsInactive
                FROM VMaterialLocationWiseBalances
                INNER JOIN ItemsRMComp ON VMaterialLocationWiseBalances.RMID = ItemsRMComp.RMID
                WHERE ItemsRMComp.ItemID = @ItemCode AND ShelfQty > 0
                ORDER BY MaterialID";
            var forgList = await db.QueryAsync<ForgingRawMaterialItem>(forgSql, new { ItemCode = itemCode });
            data.Forgings = forgList.ToList();
            data.Summary.TotalForging = data.Forgings.Sum(x => x.ShelfQty);

            // 8. Process Routings
            const string procSql = @"
                SELECT ItemSNo AS SNo, Description, ISNULL(Rate, 0) AS Rate
                FROM VItemProcesses
                WHERE IsExist = @ItemCode
                ORDER BY ItemSNo";
            var procList = await db.QueryAsync<ItemProcessRoutingItem>(procSql, new { ItemCode = itemCode });
            data.ProcessRoutings = procList.ToList();

            // 9. Maker Rates
            const string makerRateSql = @"
                SELECT Makers.VenderName AS MakerName, Processes.Description AS ProcessDescription,
                       CAST(ISNULL(VendAssItems.Rate, 0) AS varchar(20)) AS RateDisplay
                FROM VendAssItems
                INNER JOIN Makers ON VendAssItems.VendID = Makers.VendID
                INNER JOIN Processes ON VendAssItems.ProcessID = Processes.ProcessID
                WHERE ItemID = @ItemCode
                ORDER BY Processes.ProcessID";
            var makerRateList = await db.QueryAsync<MakerRateAssignmentItem>(makerRateSql, new { ItemCode = itemCode });
            data.MakerRates = makerRateList.ToList();

            // 10. Customer Catalogs
            const string catSql = @"
                SELECT FCustomerCatalog.CustCode, ItemID, ISNULL(OCR, '') AS OCR, ISNULL(TOrders.Qty, 0) AS OrderedQty
                FROM FCustomerCatalog
                LEFT OUTER JOIN (
                    SELECT CustCode, CompItemCode, SUM(Qty) AS Qty
                    FROM FOrderItems
                    INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo = FOrderItems.OrderNo
                    WHERE DT >= '2012-01-01'
                    GROUP BY CustCode, CompItemCode
                ) TOrders ON FCustomerCatalog.CustCode = TOrders.CustCode AND FCustomerCatalog.CompItemID = TOrders.CompItemCode
                WHERE CompItemID = @ItemCode AND FCustomerCatalog.CustCode <> 'Stock'";
            var catList = await db.QueryAsync<CustomerCatalogRefItem>(catSql, new { ItemCode = itemCode });
            data.CustomerCatalogs = catList.ToList();

            // 11. Customer Complaints
            const string compSql = @"
                SELECT CustomerComplaints.EntryID, CustomerComplaints.ComplaintNo,
                       ISNULL(CustomerComplaints.Description, ISNULL(CustomerComplaints_Detail.ItemReason, '')) AS Description
                FROM CustomerComplaints
                INNER JOIN CustomerComplaints_Detail ON CustomerComplaints.EntryID = CustomerComplaints_Detail.RefID
                WHERE CustomerComplaints_Detail.ItemID = @ItemCode AND CustomerComplaints.CustCode <> 'Stock'";
            var compList = await db.QueryAsync<CustomerComplaintItem>(compSql, new { ItemCode = itemCode });
            data.Complaints = compList.ToList();

            // 12. Net Planning Equation:
            // QtyRequired = CustOrders + MinStock - (RunningLots + OpenPOs + SFStock + RFStock)
            data.Summary.QtyRequired = data.Summary.TotalCustOrders + data.Summary.MinStock -
                (data.Summary.TotalRunningLots + data.Summary.TotalOpenPO + data.Summary.TotalSFStock + data.Summary.TotalRFStock);

            // 13. Weighted Milestone Completion Rate
            try
            {
                long totalProcesses = await db.ExecuteScalarAsync<long>(
                    "SELECT COUNT(*) FROM ItemProcesses WHERE ItemID = @ItemCode",
                    new { ItemCode = itemCode });

                if (totalProcesses > 0)
                {
                    long totalOrdQty = await db.ExecuteScalarAsync<long>(@"
                        SELECT ISNULL(SUM(RcvdQty), 0)
                        FROM VendRcvdDetail
                        WHERE ItemCode = @ItemCode
                          AND EntryID IN (SELECT MAX(EntryID) FROM VendRcvdDetail GROUP BY LotNo)",
                        new { ItemCode = itemCode });

                    if (totalOrdQty > 0)
                    {
                        var completedStages = await db.QueryAsync<(long SNo, long CompletedQty)>(@"
                            SELECT SNo, SUM(RcvdQty) AS CompletedQty
                            FROM VendRcvdDetail
                            INNER JOIN ItemProcesses ON VendRcvdDetail.ProcessID = ItemProcesses.ProcessID 
                                                    AND VendRcvdDetail.ItemCode = ItemProcesses.ItemID
                            WHERE ItemID = @ItemCode
                              AND VendRcvdDetail.EntryID IN (SELECT MAX(EntryID) FROM VendRcvdDetail GROUP BY LotNo)
                            GROUP BY SNo",
                            new { ItemCode = itemCode });

                        double totalCompletion = 0;
                        foreach (var stage in completedStages)
                        {
                            double stageProgress = ((double)stage.SNo / totalProcesses) * 100.0;
                            double stageWeighted = ((double)stage.CompletedQty / totalOrdQty) * stageProgress;
                            totalCompletion += stageWeighted;
                        }
                        data.Summary.CompletionPercentage = Math.Round(Math.Min(100.0, totalCompletion), 1);
                    }
                }
            }
            catch
            {
                data.Summary.CompletionPercentage = 0;
            }

            return data;
        }

        public async Task<List<ForgingCrossDemandItem>> GetForgingCrossDemandAsync(long rmid, string currentItemCode)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT VItemsRMComp.ItemID, VItemsRMComp.ItemName, ISNULL(T1.UnShippedQty, 0) AS UnShippedQty
                FROM VItemsRMComp
                INNER JOIN (
                    SELECT CompItemCode, SUM(Qty - ShippedQty) AS UnShippedQty
                    FROM VFOrderItemsWithShippedQty
                    WHERE Qty > ShippedQty
                      AND OrderNo IN (SELECT OrderNo FROM VUnShippedOrderList)
                      AND OrderNo NOT IN ('Stock-Order', 'Stock-OrderIss')
                    GROUP BY CompItemCode
                ) T1 ON VItemsRMComp.ItemID = T1.CompItemCode
                WHERE RMID = @RMID AND VItemsRMComp.ItemID <> @CurrentItemCode";

            var result = await db.QueryAsync<ForgingCrossDemandItem>(sql, new { RMID = rmid, CurrentItemCode = currentItemCode });
            return result.ToList();
        }

        public async Task<bool> ToggleForgingFunctionalStatusAsync(string itemCode, long rmid, bool isInactive)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "UPDATE ItemsRMComp SET Functional_Status = @Status WHERE RMID = @RMID AND ItemID = @ItemID";
            int rows = await db.ExecuteAsync(sql, new { Status = isInactive ? 1 : 0, RMID = rmid, ItemID = itemCode });
            return rows > 0;
        }

        public async Task<bool> CloseLotAsync(string lotNo, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "INSERT INTO Lots_Closed (LotNo, UserName, MachineName) VALUES (@LotNo, @UserName, @MachineName)";
            int rows = await db.ExecuteAsync(sql, new { LotNo = lotNo, UserName = userName, MachineName = machineName });
            return rows > 0;
        }

        public async Task<bool> ClosePOAsync(long entryID)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "UPDATE VendIssued SET Closed = 1 WHERE EntryID = @EntryID";
            int rows = await db.ExecuteAsync(sql, new { EntryID = entryID });
            return rows > 0;
        }
    }
}
