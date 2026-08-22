using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Dashboard;
using DataAccessLibrary.Models.ViewModels.Dashboard;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Dashboard
{
    public class CommandCenterDashboardDataAccess : ICommandCenterDashboardDataAccess
    {
        private readonly IConfiguration _config;

        public CommandCenterDashboardDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<CommandCenterDashboardData> GetCommandCenterDataAsync()
        {
            var data = new CommandCenterDashboardData { AsOfDate = DateTime.Now };
            using IDbConnection db = new SqlConnection(ConnectionString);

            // ── 1. HR & PAYROLL ──────────────────────────────────────────────
            try
            {
                var hrStats = await db.QueryFirstOrDefaultAsync<(int TotalActive, decimal TotalSalary)>(
                    "SELECT COUNT(EmpID) AS TotalActive, ISNULL(SUM(StartingSalary), 0) AS TotalSalary FROM Employees WHERE Active = 1");
                data.HR.TotalActiveEmployees = hrStats.TotalActive;
                data.HR.TotalMonthlyPayroll = hrStats.TotalSalary;

                int present = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(EmpID) FROM AttendanceSheet WHERE Attendance = 1 AND DT = CAST(GETDATE() AS DATE)");
                data.HR.TotalPresentEmployees = present;

                decimal longTerm = await db.ExecuteScalarAsync<decimal>(
                    "SELECT ISNULL(SUM(LongTermBalance), 0) FROM VLongTermBalance");
                data.HR.LongTermLoansBalance = longTerm;

                decimal shortTerm = await db.ExecuteScalarAsync<decimal>(
                    "SELECT ISNULL(SUM(Amount), 0) FROM Advances WHERE Type = 0 AND Cleared = 0");
                data.HR.ShortTermAdvancesBalance = shortTerm;
            }
            catch { /* Handled with default zeros */ }

            // ── 2. STOCK & INVENTORY VALUATION ───────────────────────────────
            try
            {
                var finished = await db.QueryFirstOrDefaultAsync<(long Qty, decimal Val)>(
                    "SELECT ISNULL(SUM(NetQty), 0) AS Qty, ISNULL(SUM(NetQty * PriceForCost), 0) AS Val FROM VItemsWithShelfWiseStock INNER JOIN Items ON VItemsWithShelfWiseStock.ItemID = Items.ItemID");
                data.Stock.FinishedStockQty = finished.Qty;
                data.Stock.FinishedStockValue = finished.Val;

                var sf = await db.QueryFirstOrDefaultAsync<(long Qty, decimal Val)>(@"
                    SELECT ISNULL(SUM(RcvdQty - IssQty), 0) AS Qty, ISNULL(SUM((RcvdQty - IssQty) * PriceForCost), 0) AS Val
                    FROM VendRcvdDetail
                    INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID = VStockOrderOpening.EntryID
                    INNER JOIN Items ON VendRcvdDetail.ItemCode = Items.ItemID
                    WHERE (RcvdQty - IssQty) > 0");
                data.Stock.SFStockQty = sf.Qty;
                data.Stock.SFStockValue = sf.Val;

                var forging = await db.QueryFirstOrDefaultAsync<(long Qty, decimal Val)>(
                    "SELECT ISNULL(SUM(QtyInStock), 0) AS Qty, ISNULL(SUM(QtyInStock * Rate), 0) AS Val FROM VRM WHERE GroupID IN (SELECT Group_ID FROM RMGroupIDsForForging)");
                data.Stock.ForgingStockQty = forging.Qty;
                data.Stock.ForgingStockValue = forging.Val;

                var other = await db.QueryFirstOrDefaultAsync<(long Qty, decimal Val)>(
                    "SELECT ISNULL(SUM(QtyInStock), 0) AS Qty, ISNULL(SUM(QtyInStock * Rate), 0) AS Val FROM VRM WHERE GroupID NOT IN (SELECT Group_ID FROM RMGroupIDsForForging)");
                data.Stock.OtherStockQty = other.Qty;
                data.Stock.OtherStockValue = other.Val;

                var rework = await db.QueryFirstOrDefaultAsync<(int Lots, long Qty)>(
                    "SELECT COUNT(*) AS Lots, ISNULL(SUM(Qty - IssQty), 0) AS Qty FROM VVendRcvdDetailReWorkDetail WHERE Qty > IssQty");
                data.Stock.TotalReWorkLots = rework.Lots;
                data.Stock.TotalReWorkQty = rework.Qty;
            }
            catch { /* Handled with default zeros */ }

            // ── 3. FINANCIALS & TREASURY ─────────────────────────────────────
            try
            {
                data.Financial.CashInHand = await db.ExecuteScalarAsync<decimal>(
                    "SELECT ISNULL(SUM(Balance), 0) FROM Accounts WHERE SubAccOf = '15-001'");

                data.Financial.BankBalance = await db.ExecuteScalarAsync<decimal>(
                    "SELECT ISNULL(SUM(Balance), 0) FROM Accounts WHERE SubAccOf = '15-002'");

                data.Financial.AccountsPayable = await db.ExecuteScalarAsync<decimal>(
                    "SELECT ISNULL(SUM(Balance), 0) FROM Accounts WHERE LEFT(AccNo, 6) = '24-004'");

                data.Financial.AccountsReceivable = await db.ExecuteScalarAsync<decimal>(
                    "SELECT ISNULL(SUM(Balance), 0) FROM Accounts WHERE SubAccOf = '15-003'");

                data.Financial.FixedAssetsCost = await db.ExecuteScalarAsync<decimal>(
                    "SELECT ISNULL(SUM(AcquisitionCost), 0) FROM FixedAssets");
            }
            catch { /* Handled with default zeros */ }

            // ── 4. EXPORT & DEMAND ───────────────────────────────────────────
            try
            {
                var exportInHand = await db.QueryFirstOrDefaultAsync<(long Qty, decimal Val)>(@"
                    SELECT ISNULL(SUM(Qty - ShippedQty), 0) AS Qty,
                           ISNULL(SUM((Qty - ShippedQty) * (Price * ExchRate)), 0) AS Val
                    FROM VFOrderItemsWithShippedQty
                    WHERE OrderNo IN (SELECT OrderNo FROM VUnShippedOrderList)
                      AND Qty > ShippedQty AND CustCode <> 'Stock'");
                data.Export.InHandQty = exportInHand.Qty;
                data.Export.InHandValue = exportInHand.Val;

                data.Export.LifetimeTotalSales = await db.ExecuteScalarAsync<decimal>(
                    "SELECT ISNULL(ABS(SUM(Balance)), 0) FROM Accounts WHERE SubAccOf = '31-001'");

                // Calculate Current Fiscal Year Date Range (July 1 - June 30)
                DateTime today = DateTime.Today;
                DateTime fyFrom, fyTo;
                if (today.Month >= 7)
                {
                    fyFrom = new DateTime(today.Year, 7, 1);
                    fyTo = new DateTime(today.Year + 1, 6, 30);
                }
                else
                {
                    fyFrom = new DateTime(today.Year - 1, 7, 1);
                    fyTo = new DateTime(today.Year, 6, 30);
                }

                data.Export.TotalSalesCurrentFY = await db.ExecuteScalarAsync<decimal>(@"
                    SELECT ISNULL(SUM(Credit - Debit), 0)
                    FROM Vouchers
                    WHERE LEFT(AccNo, 6) = '31-001' AND VDate BETWEEN @From AND @To",
                    new { From = fyFrom, To = fyTo });

                var monthlyTrend = await db.QueryAsync<MonthlyExportSalesTrendItem>(@"
                    SELECT TOP 12 
                        FORMAT(CustomInvoice.DT, 'MMM yy') AS MonthLabel,
                        YEAR(CustomInvoice.DT) AS [Year],
                        MONTH(CustomInvoice.DT) AS [Month],
                        COUNT(DISTINCT CustomInvoice.CustomInvoice) AS InvoiceCount,
                        ISNULL(SUM(VCustomInvoiceTotalQty.TotalQty), 0) AS TotalShippedQty,
                        ISNULL(SUM(VCustomInvoiceAmt.TotalAmt), 0) AS TotalAmountForeign,
                        ISNULL(SUM(VCustomInvoiceAmt.TotalCustomAmt), 0) AS TotalAmountPKR
                    FROM CustomInvoice
                    INNER JOIN VCustomInvoiceTotalQty ON CustomInvoice.CustomInvoice = VCustomInvoiceTotalQty.CustomInvoice
                    INNER JOIN VCustomInvoiceAmt ON CustomInvoice.CustomInvoice = VCustomInvoiceAmt.CustomInvoice
                    GROUP BY FORMAT(CustomInvoice.DT, 'MMM yy'), YEAR(CustomInvoice.DT), MONTH(CustomInvoice.DT)
                    ORDER BY YEAR(CustomInvoice.DT) ASC, MONTH(CustomInvoice.DT) ASC");

                data.Export.MonthlyTrend = monthlyTrend.ToList();
            }
            catch { /* Handled with default zeros */ }

            // ── 5. PRODUCTION THROUGHPUT & BOTTLENECK STATIONS ───────────────
            try
            {
                var forgOrders = await db.QueryFirstOrDefaultAsync<(int Count, long Qty)>(
                    "SELECT COUNT(*) AS Count, ISNULL(SUM(QtyToRcv), 0) AS Qty FROM VVendOrdersToRcv WHERE GroupID IN (SELECT Group_ID FROM RMGroupIDsForForging)");
                data.Production.ForgingOrdersCount = forgOrders.Count;
                data.Production.ForgingOrdersQty = forgOrders.Qty;

                var makerOrders = await db.QueryFirstOrDefaultAsync<(int Count, long Qty)>(@"
                    SELECT COUNT(DISTINCT MasterPONo) AS Count,
                           ISNULL(SUM(VendIssdDetail_ReturnDTs.IssQty - VendIssdDetail_ReturnDTs.RcvQty + ISNULL(WastageQty, 0)), 0) AS Qty
                    FROM VendIssued
                    INNER JOIN VendIssdDetail ON VendIssued.EntryID = VendIssdDetail.RefID
                    INNER JOIN VendIssdDetail_ReturnDTs ON VendIssdDetail.EntryID = VendIssdDetail_ReturnDTs.VIS_RefID
                    INNER JOIN Makers ON VendIssued.VendID = Makers.VendID
                    WHERE Closed = 0 AND LotNo = '0'
                      AND VendIssdDetail_ReturnDTs.IssQty + CASE WHEN ISNULL(IssRcvBalance, 0) < 0 THEN ABS(ISNULL(IssRcvBalance, 0)) ELSE 0 END + ISNULL(Wastage_Return, 0) > VendIssdDetail_ReturnDTs.RcvQty");
                data.Production.MakerOrdersCount = makerOrders.Count;
                data.Production.MakerOrdersQty = makerOrders.Qty;

                var runningLots = await db.QueryFirstOrDefaultAsync<(int Count, long Qty)>(
                    "SELECT COUNT(*) AS Count, ISNULL(SUM(Qty), 0) AS Qty FROM VRunningLots_Simple");
                data.Production.RunningLotsCount = runningLots.Count;
                data.Production.RunningLotsQty = runningLots.Qty;

                // Bottleneck Process IDs from GeneralData
                var procConfigs = (await db.QueryAsync<(string DataName, string DataValue)>(@"
                    SELECT DataName, DataValue
                    FROM GeneralData
                    WHERE DataName IN ('HeatTreatmentProcessID', 'PolishingProcessID', 'Q.CInspection100%', 'PackingProcessID')"))
                    .ToDictionary(x => x.DataName, x => long.TryParse(x.DataValue, out long v) ? v : 0);

                data.Production.HeatTreatmentProcessID = procConfigs.GetValueOrDefault("HeatTreatmentProcessID", 5);
                data.Production.PolishingProcessID = procConfigs.GetValueOrDefault("PolishingProcessID", 151);
                data.Production.QCProcessID = procConfigs.GetValueOrDefault("Q.CInspection100%", 258);
                data.Production.PackingProcessID = procConfigs.GetValueOrDefault("PackingProcessID", 138);

                // Lots at Temper
                if (data.Production.HeatTreatmentProcessID > 0)
                {
                    var temper = await db.QueryFirstOrDefaultAsync<(int Count, long Qty)>(
                        "SELECT COUNT(*), ISNULL(SUM(Qty), 0) FROM VRunningLots_Simple WHERE ProcessID = @PID",
                        new { PID = data.Production.HeatTreatmentProcessID });
                    data.Production.LotsAtTemperCount = temper.Count;
                    data.Production.LotsAtTemperQty = temper.Qty;
                }

                // Lots at Polishing
                if (data.Production.PolishingProcessID > 0)
                {
                    var polish = await db.QueryFirstOrDefaultAsync<(int Count, long Qty)>(
                        "SELECT COUNT(*), ISNULL(SUM(Qty), 0) FROM VRunningLots_Simple WHERE ProcessID = @PID",
                        new { PID = data.Production.PolishingProcessID });
                    data.Production.LotsAtPolishingCount = polish.Count;
                    data.Production.LotsAtPolishingQty = polish.Qty;
                }

                // Lots at QC
                if (data.Production.QCProcessID > 0)
                {
                    var qc = await db.QueryFirstOrDefaultAsync<(int Count, long Qty)>(
                        "SELECT COUNT(*), ISNULL(SUM(Qty), 0) FROM VRunningLots_Simple WHERE ProcessID = @PID",
                        new { PID = data.Production.QCProcessID });
                    data.Production.LotsAtQCCount = qc.Count;
                    data.Production.LotsAtQCQty = qc.Qty;
                }

                // Lots at Packed
                if (data.Production.PackingProcessID > 0)
                {
                    var packed = await db.QueryFirstOrDefaultAsync<(int Count, long Qty)>(
                        "SELECT COUNT(*), ISNULL(SUM(Qty), 0) FROM VRunningLots_Simple WHERE ProcessID = @PID",
                        new { PID = data.Production.PackingProcessID });
                    data.Production.LotsAtPackedCount = packed.Count;
                    data.Production.LotsAtPackedQty = packed.Qty;
                }
            }
            catch { /* Handled with default zeros */ }

            return data;
        }

        public async Task<List<StockGroupSummaryItem>> GetStockGroupSummaryAsync(int stockType)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql;

            if (stockType == 0) // Finished Goods
            {
                sql = @"
                    SELECT Items.GroupID, ItemGroups.Description AS GroupName,
                           ISNULL(SUM(NetQty), 0) AS TotalStock,
                           ISNULL(SUM(NetQty * PriceForCost), 0) AS TotalValue
                    FROM VItemsWithShelfWiseStock
                    INNER JOIN Items ON VItemsWithShelfWiseStock.ItemID = Items.ItemID
                    INNER JOIN ItemGroups ON Items.GroupID = ItemGroups.ID
                    GROUP BY Items.GroupID, ItemGroups.Description
                    ORDER BY TotalValue DESC";
            }
            else if (stockType == 1) // Semi-Finished Goods
            {
                sql = @"
                    SELECT Items.GroupID, ItemGroups.Description AS GroupName,
                           ISNULL(SUM(Qty - IssdQty), 0) AS TotalStock,
                           ISNULL(SUM((Qty - IssdQty) * PriceForCost), 0) AS TotalValue
                    FROM VStockOrderOpening
                    INNER JOIN (SELECT Opening_RefID FROM VendRcvdDetail GROUP BY Opening_RefID) T1 ON T1.Opening_RefID = VStockOrderOpening.EntryID
                    INNER JOIN Items ON VStockOrderOpening.ItemID = Items.ItemID
                    INNER JOIN ItemGroups ON Items.GroupID = ItemGroups.ID
                    WHERE Qty > IssdQty
                    GROUP BY Items.GroupID, ItemGroups.Description
                    ORDER BY TotalValue DESC";
            }
            else if (stockType == 2) // Forgings
            {
                sql = @"
                    SELECT VRM.GroupID, VRM.GroupName AS GroupName,
                           ISNULL(SUM(QtyInStock), 0) AS TotalStock,
                           ISNULL(SUM(QtyInStock * Rate), 0) AS TotalValue
                    FROM VRM
                    WHERE GroupID IN (SELECT Group_ID FROM RMGroupIDsForForging) AND QtyInStock > 0
                    GROUP BY VRM.GroupID, VRM.GroupName
                    ORDER BY TotalValue DESC";
            }
            else // Other Inventory
            {
                sql = @"
                    SELECT VRM.GroupID, VRM.GroupName AS GroupName,
                           ISNULL(SUM(QtyInStock), 0) AS TotalStock,
                           ISNULL(SUM(QtyInStock * Rate), 0) AS TotalValue
                    FROM VRM
                    WHERE GroupID NOT IN (SELECT Group_ID FROM RMGroupIDsForForging) AND QtyInStock > 0
                    GROUP BY VRM.GroupID, VRM.GroupName
                    ORDER BY TotalValue DESC";
            }

            var result = await db.QueryAsync<StockGroupSummaryItem>(sql);
            return result.ToList();
        }

        public async Task<List<StockGroupItemDetail>> GetStockGroupDetailAsync(int stockType, int groupID)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql;

            if (stockType == 0) // Finished Goods
            {
                sql = @"
                    SELECT Items.ItemID, Items.ItemName,
                           ISNULL(SUM(NetQty), 0) AS TotalStock,
                           ISNULL(SUM(NetQty * PriceForCost), 0) AS TotalValue
                    FROM VItemsWithShelfWiseStock
                    INNER JOIN Items ON VItemsWithShelfWiseStock.ItemID = Items.ItemID
                    WHERE Items.GroupID = @GroupID
                    GROUP BY Items.ItemID, Items.ItemName
                    ORDER BY TotalValue DESC";
            }
            else if (stockType == 1) // Semi-Finished Goods
            {
                sql = @"
                    SELECT Items.ItemID, Items.ItemName,
                           ISNULL(SUM(Qty - IssdQty), 0) AS TotalStock,
                           ISNULL(SUM((Qty - IssdQty) * PriceForCost), 0) AS TotalValue
                    FROM VStockOrderOpening
                    INNER JOIN (SELECT Opening_RefID FROM VendRcvdDetail GROUP BY Opening_RefID) T1 ON T1.Opening_RefID = VStockOrderOpening.EntryID
                    INNER JOIN Items ON VStockOrderOpening.ItemID = Items.ItemID
                    WHERE Items.GroupID = @GroupID AND Qty > IssdQty
                    GROUP BY Items.ItemID, Items.ItemName
                    ORDER BY TotalValue DESC";
            }
            else if (stockType == 2 || stockType == 3) // Forgings or Other RM
            {
                sql = @"
                    SELECT VRM.RMID1 AS ItemID, VRM.RMName AS ItemName,
                           ISNULL(SUM(QtyInStock), 0) AS TotalStock,
                           ISNULL(SUM(QtyInStock * Rate), 0) AS TotalValue
                    FROM VRM
                    WHERE VRM.GroupID = @GroupID AND QtyInStock > 0
                    GROUP BY VRM.RMID1, VRM.RMName
                    ORDER BY TotalValue DESC";
            }
            else
            {
                return new List<StockGroupItemDetail>();
            }

            var result = await db.QueryAsync<StockGroupItemDetail>(sql, new { GroupID = groupID });
            return result.ToList();
        }
    }
}
