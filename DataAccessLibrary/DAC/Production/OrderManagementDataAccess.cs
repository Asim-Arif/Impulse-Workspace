using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Production
{
    public class OrderManagementDataAccess : IOrderManagementDataAccess
    {
        private readonly IConfiguration _config;

        public OrderManagementDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("ImpulseConnection")
            ?? string.Empty;

        public async Task<List<LookupItemString>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT DISTINCT CustCode AS Id, CustCode AS Name 
                FROM ForeignCustomers 
                WHERE CustCode IS NOT NULL AND CustCode <> ''
                ORDER BY CustCode";
            var list = (await db.QueryAsync<LookupItemString>(sql)).ToList();
            list.Insert(0, new LookupItemString { Id = "", Name = "<All Customers>" });
            return list;
        }

        public async Task<List<CustomerOrderHeaderDto>> GetOrdersAsync(OrderManagementFilter filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            var sb = new StringBuilder(@"
                SELECT 
                    ol.OrderNo,
                    ol.InternalRefNo,
                    ol.DT,
                    ol.CustCode,
                    ol.Country,
                    ol.DeliveryDT,
                    ISNULL(ol.TotalOrderQty, 0) AS TotalOrderQty,
                    ISNULL(ol.TotalPlannedQty, 0) AS TotalPlannedQty,
                    ISNULL(ol.TotalShippedQty, 0) AS TotalShippedQty,
                    ISNULL(ol.TotalArticles, 0) AS TotalArticles,
                    ISNULL(ol.Authorized, 0) AS Authorized,
                    ISNULL(ol.OrderPlanApproved, 0) AS OrderPlanApproved
                FROM VFOrderList ol
                WHERE 1 = 1");

            var p = new DynamicParameters();

            if (!string.IsNullOrWhiteSpace(filter.CustCode))
            {
                sb.Append(" AND ol.CustCode = @CustCode");
                p.Add("@CustCode", filter.CustCode);
            }

            if (!string.IsNullOrWhiteSpace(filter.SearchText))
            {
                sb.Append(" AND (ol.OrderNo LIKE @Search OR ol.InternalRefNo LIKE @Search)");
                p.Add("@Search", $"%{filter.SearchText.Trim()}%");
            }

            // Date filtering
            switch (filter.DateRangeType)
            {
                case 1: // Last 30 Days
                    sb.Append(" AND ol.DT >= DATEADD(day, -30, GETDATE())");
                    break;
                case 2: // Last 90 Days
                    sb.Append(" AND ol.DT >= DATEADD(day, -90, GETDATE())");
                    break;
                case 3: // This Year
                    sb.Append(" AND YEAR(ol.DT) = YEAR(GETDATE())");
                    break;
                case 4: // Custom Range
                    if (filter.FromDate.HasValue)
                    {
                        sb.Append(" AND ol.DT >= @FromDate");
                        p.Add("@FromDate", filter.FromDate.Value.Date);
                    }
                    if (filter.ToDate.HasValue)
                    {
                        sb.Append(" AND ol.DT <= @ToDate");
                        p.Add("@ToDate", filter.ToDate.Value.Date.AddDays(1).AddTicks(-1));
                    }
                    break;
            }

            sb.Append(" ORDER BY ol.DT DESC");

            var orders = (await db.QueryAsync<CustomerOrderHeaderDto>(sb.ToString(), p)).ToList();

            if (!orders.Any())
                return orders;

            // Fetch high-level stage info for these orders
            var orderNos = orders.Select(x => x.OrderNo).Take(200).ToList();

            // 1. Shipped / Invoiced quantities
            var shippedMap = (await db.QueryAsync<(string OrderNo, int ShippedQty)>(@"
                SELECT OrderNo, SUM(ShippedQty) AS ShippedQty
                FROM VFOrderItems
                WHERE OrderNo IN @OrderNos
                GROUP BY OrderNo", new { OrderNos = orderNos }))
                .ToDictionary(x => x.OrderNo, x => x.ShippedQty, StringComparer.OrdinalIgnoreCase);

            // 2. Dispatched quantities
            var dispatchedMap = (await db.QueryAsync<(string OrderNo, int DispatchedQty)>(@"
                SELECT vrd.OrderNo, ISNULL(SUM(dvrd.Qty), 0) AS DispatchedQty
                FROM DispatchListDetail_VRD dvrd
                INNER JOIN VendRcvdDetail vrd ON dvrd.VRD_RefID = vrd.EntryID
                WHERE vrd.OrderNo IN @OrderNos
                GROUP BY vrd.OrderNo", new { OrderNos = orderNos }))
                .ToDictionary(x => x.OrderNo, x => x.DispatchedQty, StringComparer.OrdinalIgnoreCase);

            // 3. Produced / Final QC quantities
            var producedMap = (await db.QueryAsync<(string OrderNo, int ProducedQty)>(@"
                SELECT vrd.OrderNo, ISNULL(SUM(vrd.RcvdQty), 0) AS ProducedQty
                FROM VendRcvdDetail vrd
                INNER JOIN Processes p ON vrd.ProcessID = p.ProcessID
                WHERE vrd.OrderNo IN @OrderNos 
                  AND (p.Description LIKE '%Ready-Finish%' OR p.Description LIKE '%Job Card Close%' OR p.Description LIKE '%Final QC%')
                GROUP BY vrd.OrderNo", new { OrderNos = orderNos }))
                .ToDictionary(x => x.OrderNo, x => x.ProducedQty, StringComparer.OrdinalIgnoreCase);

            // 4. Running lots details for orders
            var runningLots = (await db.QueryAsync<dynamic>(@"
                SELECT 
                    rl.OrderNo,
                    ISNULL(rl.Qty, 0) AS Qty,
                    COALESCE(currPgp.SeqNo, CAST(rl.SNO AS INT), 1) AS CurrentSeqNo,
                    COALESCE(totPgp.TotalSteps, CAST(rl.MaxSno AS INT), 1) AS TotalSeqNo
                FROM VRunningLots_Simple rl
                LEFT JOIN ItemProcessGroups ipg ON rl.ItemCode = ipg.ItemID
                LEFT JOIN ProcessGroupsProcesses currPgp ON ipg.PG_RefID = currPgp.Group_RefID AND rl.ProcessID = currPgp.Process_RefID
                LEFT JOIN (
                    SELECT Group_RefID, COUNT(1) AS TotalSteps
                    FROM ProcessGroupsProcesses
                    GROUP BY Group_RefID
                ) totPgp ON ipg.PG_RefID = totPgp.Group_RefID
                WHERE rl.OrderNo IN @OrderNos", new { OrderNos = orderNos }))
                .GroupBy(x => (string)x.OrderNo, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

            // 5. PO Issuance & Receiving summary
            var poMap = (await db.QueryAsync<(string OrderNo, int TotalIssQty, int TotalRcvdQty)>(@"
                SELECT 
                    vid.OrderNo,
                    ISNULL(SUM(vid.IssQty), 0) AS TotalIssQty,
                    ISNULL(SUM(ISNULL(rcv.TotalRcvd, ISNULL(vid.RcvdQty, 0))), 0) AS TotalRcvdQty
                FROM VendIssdDetail vid
                INNER JOIN VendIssued vi ON vid.RefID = vi.EntryID
                LEFT JOIN (
                    SELECT vr.Issuance_RefID, vrd.ItemCode, SUM(vrd.RcvdQty) AS TotalRcvd
                    FROM VendReceived vr
                    INNER JOIN VendRcvdDetail vrd ON vr.EntryID = vrd.RefID
                    GROUP BY vr.Issuance_RefID, vrd.ItemCode
                ) rcv ON vi.EntryID = rcv.Issuance_RefID AND vid.ItemCode = rcv.ItemCode
                WHERE vid.OrderNo IN @OrderNos
                  AND ISNULL(vi.MasterPONo, '') <> ''
                GROUP BY vid.OrderNo", new { OrderNos = orderNos }))
                .ToDictionary(x => x.OrderNo, x => (x.TotalIssQty, x.TotalRcvdQty), StringComparer.OrdinalIgnoreCase);

            foreach (var ord in orders)
            {
                int orderQty = Math.Max(1, ord.TotalOrderQty);
                shippedMap.TryGetValue(ord.OrderNo, out int shippedQty);
                dispatchedMap.TryGetValue(ord.OrderNo, out int dispQty);
                producedMap.TryGetValue(ord.OrderNo, out int prodQty);
                runningLots.TryGetValue(ord.OrderNo, out var ordLots);
                ordLots ??= new List<dynamic>();
                poMap.TryGetValue(ord.OrderNo, out var poInfo);

                bool hasRunningLots = ordLots.Any();
                bool hasPos = poInfo.TotalIssQty > 0;

                // 1. Stage determination
                if (shippedQty >= orderQty && orderQty > 0)
                {
                    ord.CurrentStageNumber = 10;
                    ord.CurrentStageName = "Invoiced & Shipped";
                }
                else if (dispQty >= orderQty && orderQty > 0)
                {
                    ord.CurrentStageNumber = 9;
                    ord.CurrentStageName = "Dispatch Finalized";
                }
                else if (dispQty > 0)
                {
                    ord.CurrentStageNumber = 8;
                    ord.CurrentStageName = "Dispatch Staging";
                }
                else if (prodQty >= orderQty && orderQty > 0)
                {
                    ord.CurrentStageNumber = 7;
                    ord.CurrentStageName = "Production Complete";
                }
                else if (prodQty > 0)
                {
                    ord.CurrentStageNumber = 7;
                    ord.CurrentStageName = "Partial Production";
                }
                else if (hasRunningLots)
                {
                    ord.CurrentStageNumber = 6;
                    ord.CurrentStageName = "Floor Production";
                }
                else if (hasPos)
                {
                    ord.CurrentStageNumber = 4;
                    ord.CurrentStageName = "PO Issued";
                }
                else if (ord.OrderPlanApproved || ord.TotalPlannedQty > 0)
                {
                    ord.CurrentStageNumber = 3;
                    ord.CurrentStageName = "PPC Planned";
                }
                else if (ord.Authorized)
                {
                    ord.CurrentStageNumber = 2;
                    ord.CurrentStageName = "Authorized";
                }
                else
                {
                    ord.CurrentStageNumber = 1;
                    ord.CurrentStageName = "Order Entered";
                }

                // 2. Physical Completion Percentage
                if (ord.CurrentStageNumber == 10)
                {
                    ord.ProgressPct = 100.0;
                }
                else
                {
                    double completedUnits = 0;

                    // Shipped / Invoiced (100% complete)
                    completedUnits += Math.Min(orderQty, shippedQty) * 1.0;

                    // Staged for dispatch (95% complete)
                    int stagedDispatched = Math.Max(0, Math.Min(orderQty, dispQty) - shippedQty);
                    completedUnits += stagedDispatched * 0.95;

                    // Finished goods in store awaiting dispatch (90% complete)
                    int readyInStore = Math.Max(0, Math.Min(orderQty, prodQty) - Math.Max(dispQty, shippedQty));
                    completedUnits += readyInStore * 0.90;

                    // Running lots progressing through factory processes
                    foreach (var lot in ordLots)
                    {
                        int lotQty = (int)(lot.Qty ?? 0);
                        int curSeq = (int)(lot.CurrentSeqNo ?? 1);
                        int totSeq = (int)(lot.TotalSeqNo ?? 1);
                        double lotProgressFraction = totSeq > 0 ? Math.Min(1.0, (double)curSeq / totSeq) : 0.1;
                        completedUnits += lotQty * (lotProgressFraction * 0.90);
                    }

                    // Purchase orders issued (material procurement stage ~5%)
                    int unrcvdPo = Math.Max(0, poInfo.TotalIssQty - poInfo.TotalRcvdQty);
                    completedUnits += unrcvdPo * 0.05;

                    double calculatedPct = Math.Round((completedUnits * 100.0) / orderQty, 1);
                    ord.ProgressPct = Math.Min(99.0, Math.Max(0.0, calculatedPct));
                }
            }

            // Status post-filtering if requested
            if (filter.StatusFilter == "InProgress")
                orders = orders.Where(x => x.ProgressPct > 0 && x.ProgressPct < 100).ToList();
            else if (filter.StatusFilter == "Completed")
                orders = orders.Where(x => x.ProgressPct >= 100).ToList();
            else if (filter.StatusFilter == "Overdue")
                orders = orders.Where(x => x.IsOverdue).ToList();

            return orders;
        }

        public async Task<OrderSummaryCardDto> GetOrderSummaryMetricsAsync(OrderManagementFilter filter)
        {
            var orders = await GetOrdersAsync(filter);

            return new OrderSummaryCardDto
            {
                TotalOrders = orders.Count,
                InProgressOrders = orders.Count(x => x.ProgressPct > 0 && x.ProgressPct < 100),
                CompletedOrders = orders.Count(x => x.ProgressPct >= 100),
                OverdueOrders = orders.Count(x => x.IsOverdue),
                TotalOrderedPcs = orders.Sum(x => x.TotalOrderQty),
                TotalProducedPcs = orders.Sum(x => (int)(x.TotalOrderQty * (x.ProgressPct / 100.0))),
                TotalDispatchedPcs = orders.Sum(x => x.TotalShippedQty),
                AverageProgressPct = orders.Any() ? Math.Round(orders.Average(x => x.ProgressPct), 1) : 0
            };
        }

        public async Task<List<OrderItemProgressDto>> GetOrderItemsAsync(string orderNo)
        {
            if (string.IsNullOrWhiteSpace(orderNo))
                return new List<OrderItemProgressDto>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            const string itemsSql = @"
                SELECT 
                    oi.ID,
                    oi.OrderNo,
                    oi.ItemCode,
                    oi.CompItemCode,
                    oi.ItemName,
                    oi.Description,
                    ISNULL(oi.Qty, 0) AS OrderedQty,
                    ISNULL(oi.ShippedQty, 0) AS DispatchedQty,
                    oi.DeliveryDT,
                    ISNULL(oi.Packaging, '') AS Packaging,
                    ISNULL(oi.Quality, '') AS Quality,
                    ISNULL(oi.GroupID, 0) AS GroupID
                FROM VFOrderItems oi
                WHERE oi.OrderNo = @OrderNo
                ORDER BY oi.ID ASC";

            var items = (await db.QueryAsync<OrderItemProgressDto>(itemsSql, new { OrderNo = orderNo })).ToList();

            if (!items.Any()) return items;

            // Enrich with ProducedQty, PlannedQty, Lots, and Stage
            foreach (var item in items)
            {
                // Dispatched from VRD
                int vrdDispatched = await db.ExecuteScalarAsync<int?>(@"
                    SELECT ISNULL(SUM(dvrd.Qty), 0)
                    FROM DispatchListDetail_VRD dvrd
                    INNER JOIN VendRcvdDetail vrd ON dvrd.VRD_RefID = vrd.EntryID
                    WHERE (vrd.OrderNo = @OrderNo OR vrd.OrderNo = @CleanOrderNo)
                      AND (vrd.ItemCode = @ItemCode OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND vrd.ItemCode = @CompItemCode))",
                    new { 
                        OrderNo = orderNo.Trim(), 
                        CleanOrderNo = orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim(),
                        ItemCode = item.ItemCode?.Trim(), 
                        CompItemCode = item.CompItemCode?.Trim() 
                    }) ?? 0;

                // Finished / Ready in production
                int vrdProduced = await db.ExecuteScalarAsync<int?>(@"
                    SELECT ISNULL(SUM(vrd.RcvdQty), 0)
                    FROM VendRcvdDetail vrd
                    INNER JOIN Processes p ON vrd.ProcessID = p.ProcessID
                    WHERE (vrd.OrderNo = @OrderNo OR vrd.OrderNo = @CleanOrderNo)
                      AND (vrd.ItemCode = @ItemCode OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND vrd.ItemCode = @CompItemCode))
                      AND (p.Description LIKE '%Ready-Finish%' OR p.Description LIKE '%Job Card Close%' OR p.Description LIKE '%Final QC%')",
                    new { 
                        OrderNo = orderNo.Trim(), 
                        CleanOrderNo = orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim(),
                        ItemCode = item.ItemCode?.Trim(), 
                        CompItemCode = item.CompItemCode?.Trim() 
                    }) ?? 0;

                item.ProducedQty = Math.Max(vrdDispatched, vrdProduced);
                if (item.DispatchedQty == 0 && vrdDispatched > 0)
                    item.DispatchedQty = vrdDispatched;

                // Running lots progressing through factory processes
                var itemLots = (await db.QueryAsync<dynamic>(@"
                    SELECT 
                        rl.LotNo,
                        ISNULL(rl.Qty, 0) AS Qty,
                        COALESCE(currPgp.SeqNo, CAST(rl.SNO AS INT), 1) AS CurrentSeqNo,
                        COALESCE(totPgp.TotalSteps, CAST(rl.MaxSno AS INT), 1) AS TotalSeqNo
                    FROM VRunningLots_Simple rl
                    LEFT JOIN ItemProcessGroups ipg ON rl.ItemCode = ipg.ItemID
                    LEFT JOIN ProcessGroupsProcesses currPgp ON ipg.PG_RefID = currPgp.Group_RefID AND rl.ProcessID = currPgp.Process_RefID
                    LEFT JOIN (
                        SELECT Group_RefID, COUNT(1) AS TotalSteps
                        FROM ProcessGroupsProcesses
                        GROUP BY Group_RefID
                    ) totPgp ON ipg.PG_RefID = totPgp.Group_RefID
                    WHERE (rl.OrderNo = @OrderNo OR rl.OrderNo = @CleanOrderNo)
                      AND (rl.ItemCode = @ItemCode OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND rl.ItemCode = @CompItemCode))",
                    new { 
                        OrderNo = orderNo.Trim(), 
                        CleanOrderNo = orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim(),
                        ItemCode = item.ItemCode?.Trim(), 
                        CompItemCode = item.CompItemCode?.Trim() 
                    })).ToList();

                var activeLotNos = itemLots.Select(x => (string)x.LotNo).Where(x => !string.IsNullOrEmpty(x)).Take(2).ToList();
                item.BatchNo = activeLotNos.Any() ? string.Join(", ", activeLotNos) : (!string.IsNullOrEmpty(item.CompItemCode) ? item.CompItemCode : item.ItemCode);

                // PPC Planning (Stock vs PO vs Production)
                var ppcPlan = await db.QueryFirstOrDefaultAsync<dynamic>(@"
                    SELECT TOP 1 StockQty, TotalPurchaseQty, ProductionQty
                    FROM PPC_Order_Item_Planning
                    WHERE (OrderNo = @OrderNo OR OrderNo = @CleanOrderNo)
                      AND (ItemID = @ItemCode OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND ItemID = @CompItemCode))",
                    new { 
                        OrderNo = orderNo.Trim(), 
                        CleanOrderNo = orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim(),
                        ItemCode = item.ItemCode?.Trim(), 
                        CompItemCode = item.CompItemCode?.Trim() 
                    });

                if (ppcPlan != null)
                {
                    item.PlannedStockQty = (int)(ppcPlan.StockQty ?? 0);
                    item.PlannedPurchaseQty = (int)(ppcPlan.TotalPurchaseQty ?? 0);
                    item.PlannedProductionQty = (int)(ppcPlan.ProductionQty ?? 0);
                }

                // Physical Stock Adjustment / Allocation against order
                item.StockAdjustedQty = await db.ExecuteScalarAsync<int?>(@"
                    SELECT ISNULL(SUM(Qty), 0)
                    FROM StockOrderAdjustments
                    WHERE (OrderNo = @OrderNo OR OrderNo = @CleanOrderNo)
                      AND (ItemID = @ItemCode OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND ItemID = @CompItemCode))",
                    new { 
                        OrderNo = orderNo.Trim(), 
                        CleanOrderNo = orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim(),
                        ItemCode = item.ItemCode?.Trim(), 
                        CompItemCode = item.CompItemCode?.Trim() 
                    }) ?? 0;

                // Purchase Order Issuance and Receiving totals
                var poTotals = await db.QueryFirstOrDefaultAsync<dynamic>(@"
                    SELECT 
                        COUNT(1) AS PoCount,
                        ISNULL(SUM(vid.IssQty), 0) AS TotalIssQty,
                        ISNULL(SUM(ISNULL(rcv.TotalRcvd, ISNULL(vid.RcvdQty, 0))), 0) AS TotalRcvdQty
                    FROM VendIssdDetail vid
                    INNER JOIN VendIssued vi ON vid.RefID = vi.EntryID
                    LEFT JOIN (
                        SELECT vr.Issuance_RefID, vrd.ItemCode, SUM(vrd.RcvdQty) AS TotalRcvd
                        FROM VendReceived vr
                        INNER JOIN VendRcvdDetail vrd ON vr.EntryID = vrd.RefID
                        GROUP BY vr.Issuance_RefID, vrd.ItemCode
                    ) rcv ON vi.EntryID = rcv.Issuance_RefID AND vid.ItemCode = rcv.ItemCode
                    WHERE (vid.OrderNo = @OrderNo OR vid.OrderNo = @CleanOrderNo)
                      AND (vid.ItemCode = @ItemCode OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND vid.ItemCode = @CompItemCode))
                      AND ISNULL(vi.MasterPONo, '') <> ''",
                    new { 
                        OrderNo = orderNo.Trim(), 
                        CleanOrderNo = orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim(),
                        ItemCode = item.ItemCode?.Trim(), 
                        CompItemCode = item.CompItemCode?.Trim() 
                    });

                int poCount = (int)(poTotals?.PoCount ?? 0);
                int poIssQty = (int)(poTotals?.TotalIssQty ?? 0);
                int poRcvdQty = (int)(poTotals?.TotalRcvdQty ?? 0);

                int targetQty = Math.Max(1, item.OrderedQty);

                // Stage determination
                if (item.DispatchedQty >= targetQty && targetQty > 0)
                {
                    item.CurrentStageNumber = 10;
                    item.CurrentStageName = "Completed / Invoiced";
                }
                else if (item.DispatchedQty > 0)
                {
                    item.CurrentStageNumber = 8;
                    item.CurrentStageName = "Dispatch Staging";
                }
                else if (item.ProducedQty >= targetQty && targetQty > 0)
                {
                    item.CurrentStageNumber = 7;
                    item.CurrentStageName = "Production Complete";
                }
                else if (item.ProducedQty > 0)
                {
                    item.CurrentStageNumber = 7;
                    item.CurrentStageName = "Partial Production";
                }
                else if (itemLots.Any())
                {
                    item.CurrentStageNumber = 6;
                    item.CurrentStageName = "Floor Production";
                }
                else if (poCount > 0)
                {
                    item.CurrentStageNumber = 4;
                    item.CurrentStageName = "PO Issued";
                }
                else if (item.HasPpcPlan)
                {
                    item.CurrentStageNumber = 3;
                    item.CurrentStageName = "PPC Planned";
                }
                else
                {
                    item.CurrentStageNumber = 1;
                    item.CurrentStageName = "Order Entered";
                }

                // Physical Completion Percentage
                if (item.CurrentStageNumber == 10)
                {
                    item.ProgressPct = 100.0;
                }
                else
                {
                    double completedUnits = 0;

                    // 1. Dispatched units (100% complete)
                    completedUnits += Math.Min(targetQty, item.DispatchedQty) * 1.0;

                    // 2. Finished in store ready for dispatch (90% complete)
                    int readyInStore = Math.Max(0, Math.Min(targetQty, item.ProducedQty) - item.DispatchedQty);
                    completedUnits += readyInStore * 0.90;

                    // 3. Units progressing through factory processes in running lots
                    foreach (var lot in itemLots)
                    {
                        int lotQty = (int)(lot.Qty ?? 0);
                        int curSeq = (int)(lot.CurrentSeqNo ?? 1);
                        int totSeq = (int)(lot.TotalSeqNo ?? 1);
                        double lotProgressFraction = totSeq > 0 ? Math.Min(1.0, (double)curSeq / totSeq) : 0.1;
                        completedUnits += lotQty * (lotProgressFraction * 0.90);
                    }

                    // 4. Units in purchase orders (material procurement stage ~5%)
                    int unrcvdPo = Math.Max(0, poIssQty - poRcvdQty);
                    completedUnits += unrcvdPo * 0.05;

                    double calculatedPct = Math.Round((completedUnits * 100.0) / targetQty, 1);
                    item.ProgressPct = Math.Min(99.0, Math.Max(0.0, calculatedPct));
                }
            }

            return items;
        }

        public async Task<List<ItemPurchaseOrderDto>> GetItemPurchaseOrdersAsync(string orderNo, string itemCode, string? compItemCode = null)
        {
            if (string.IsNullOrWhiteSpace(orderNo))
                return new List<ItemPurchaseOrderDto>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            var sb = new StringBuilder(@"
                SELECT 
                    vid.EntryID,
                    vi.RecieptID AS POReceiptID,
                    ISNULL(vi.MasterPONo, '') AS MasterPONo,
                    vi.VendID,
                    ISNULL(m.VenderName, 'Unknown Supplier') AS VenderName,
                    vi.DT,
                    vi.ProcessID,
                    ISNULL(p.Description, 'Manufacturing Process') AS ProcessName,
                    ISNULL(vid.IssQty, 0) AS IssQty,
                    ISNULL(rcv.TotalRcvd, ISNULL(vid.RcvdQty, 0)) AS RcvdQty,
                    vid.ReturnDT
                FROM VendIssdDetail vid
                INNER JOIN VendIssued vi ON vid.RefID = vi.EntryID
                LEFT JOIN (
                    SELECT vr.Issuance_RefID, vrd.ItemCode, SUM(vrd.RcvdQty) AS TotalRcvd
                    FROM VendReceived vr
                    INNER JOIN VendRcvdDetail vrd ON vr.EntryID = vrd.RefID
                    GROUP BY vr.Issuance_RefID, vrd.ItemCode
                ) rcv ON vi.EntryID = rcv.Issuance_RefID AND vid.ItemCode = rcv.ItemCode
                LEFT JOIN Makers m ON vi.VendID = m.VendID
                LEFT JOIN Processes p ON vi.ProcessID = p.ProcessID
                WHERE (vid.OrderNo = @OrderNo OR vid.OrderNo = @CleanOrderNo)
                  AND ISNULL(vi.MasterPONo, '') <> ''");

            var p = new DynamicParameters();
            p.Add("@OrderNo", orderNo.Trim());
            p.Add("@CleanOrderNo", orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim());

            if (!string.IsNullOrWhiteSpace(itemCode))
            {
                sb.Append(@" AND (
                    vid.ItemCode = @ItemCode 
                    OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND vid.ItemCode = @CompItemCode)
                )");
                p.Add("@ItemCode", itemCode.Trim());
                p.Add("@CompItemCode", compItemCode?.Trim());
            }

            sb.Append(" ORDER BY vi.DT DESC, vid.EntryID DESC");

            return (await db.QueryAsync<ItemPurchaseOrderDto>(sb.ToString(), p)).ToList();
        }

        public async Task<List<ItemRunningLotDto>> GetItemRunningLotsAsync(string orderNo, string itemCode, string? compItemCode = null)
        {
            if (string.IsNullOrWhiteSpace(orderNo))
                return new List<ItemRunningLotDto>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            var sb = new StringBuilder(@"
                SELECT 
                    rl.LotNo,
                    rl.ItemCode,
                    rl.OrderNo,
                    rl.ProcessID,
                    ISNULL(p.Description, rl.Description) AS ProcessName,
                    ISNULL(rl.Qty, 0) AS Qty,
                    COALESCE(currPgp.SeqNo, CAST(rl.SNO AS INT), 1) AS CurrentSeqNo,
                    COALESCE(totPgp.TotalSteps, CAST(rl.MaxSno AS INT), 1) AS TotalSeqNo,
                    rl.DT AS LastActivityDT,
                    ISNULL(m.VenderName, 'In-House Production') AS MakerName,
                    ISNULL(rl.ReWorkLot, 0) AS ReWorkLot
                FROM VRunningLots_Simple rl
                LEFT JOIN Processes p ON rl.ProcessID = p.ProcessID
                LEFT JOIN Makers m ON rl.VendID = m.VendID
                LEFT JOIN ItemProcessGroups ipg ON rl.ItemCode = ipg.ItemID
                LEFT JOIN ProcessGroupsProcesses currPgp ON ipg.PG_RefID = currPgp.Group_RefID AND rl.ProcessID = currPgp.Process_RefID
                LEFT JOIN (
                    SELECT Group_RefID, COUNT(1) AS TotalSteps
                    FROM ProcessGroupsProcesses
                    GROUP BY Group_RefID
                ) totPgp ON ipg.PG_RefID = totPgp.Group_RefID
                WHERE (rl.OrderNo = @OrderNo OR rl.OrderNo = @CleanOrderNo)");

            var p = new DynamicParameters();
            p.Add("@OrderNo", orderNo.Trim());
            p.Add("@CleanOrderNo", orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim());

            if (!string.IsNullOrWhiteSpace(itemCode))
            {
                sb.Append(@" AND (
                    rl.ItemCode = @ItemCode 
                    OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND rl.ItemCode = @CompItemCode)
                )");
                p.Add("@ItemCode", itemCode.Trim());
                p.Add("@CompItemCode", compItemCode?.Trim());
            }

            sb.Append(" ORDER BY rl.SNO DESC, rl.LotNo ASC");

            return (await db.QueryAsync<ItemRunningLotDto>(sb.ToString(), p)).ToList();
        }

        public async Task<List<ItemDispatchDetailDto>> GetItemDispatchDetailsAsync(string orderNo, string itemCode, string? compItemCode = null)
        {
            if (string.IsNullOrWhiteSpace(orderNo))
                return new List<ItemDispatchDetailDto>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            var sb = new StringBuilder(@"
                SELECT 
                    dl.EntryID AS DispatchID,
                    dl.DispatchListNo,
                    ISNULL(dli.CartonNo, 0) AS CartonNo,
                    ISNULL(dli.InnerNo, 0) AS InnerNo,
                    vrd.LotNo,
                    ISNULL(dvrd.Qty, 0) AS Qty,
                    dvrd.EntryDT,
                    ISNULL(dvrd.AddedBy, dl.UserName) AS AddedBy,
                    ISNULL(dl.Finalyzed, 0) AS Finalyzed,
                    dl.FinalyzedDT,
                    COALESCE(invLot.CustomInvoice, invDL.CustomInvoice) AS InvoiceNo,
                    COALESCE(invLot.InvoiceDate, invDL.InvoiceDate) AS InvoiceDate
                FROM DispatchListDetail_VRD dvrd
                INNER JOIN VendRcvdDetail vrd ON dvrd.VRD_RefID = vrd.EntryID
                INNER JOIN DispatchListDetail_Inners dli ON dvrd.DLDC_RefID = dli.EntryID
                INNER JOIN DispatchListDetails_Adv dld ON dvrd.DLD_RefID = dld.EntryID
                INNER JOIN DispatchList dl ON dld.RefID = dl.EntryID
                LEFT JOIN (
                    SELECT DISTINCT 
                        cpdd.DP_RefID,
                        cp.LotNo,
                        cp.CustomInvoice,
                        ci.DT AS InvoiceDate
                    FROM CustomPList_DispatchListDetail cpdd
                    INNER JOIN CustomPList cp ON cpdd.CustomPList_RefID = cp.ID
                    INNER JOIN CustomInvoice ci ON cp.CustomInvoice = ci.CustomInvoice
                    WHERE ISNULL(cp.LotNo, '') <> ''
                ) invLot ON dl.EntryID = invLot.DP_RefID AND vrd.LotNo = invLot.LotNo
                LEFT JOIN (
                    SELECT 
                        cpdd.DP_RefID,
                        MAX(cp.CustomInvoice) AS CustomInvoice,
                        MAX(ci.DT) AS InvoiceDate
                    FROM CustomPList_DispatchListDetail cpdd
                    INNER JOIN CustomPList cp ON cpdd.CustomPList_RefID = cp.ID
                    INNER JOIN CustomInvoice ci ON cp.CustomInvoice = ci.CustomInvoice
                    GROUP BY cpdd.DP_RefID
                ) invDL ON dl.EntryID = invDL.DP_RefID
                WHERE (vrd.OrderNo = @OrderNo OR vrd.OrderNo = @CleanOrderNo)");

            var p = new DynamicParameters();
            p.Add("@OrderNo", orderNo.Trim());
            p.Add("@CleanOrderNo", orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim());

            if (!string.IsNullOrWhiteSpace(itemCode))
            {
                sb.Append(@" AND (
                    vrd.ItemCode = @ItemCode 
                    OR dld.ItemCode = @ItemCode
                    OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND (vrd.ItemCode = @CompItemCode OR dld.ItemCode = @CompItemCode))
                )");
                p.Add("@ItemCode", itemCode.Trim());
                p.Add("@CompItemCode", compItemCode?.Trim());
            }

            sb.Append(" ORDER BY dl.EntryID DESC, dli.CartonNo ASC");

            return (await db.QueryAsync<ItemDispatchDetailDto>(sb.ToString(), p)).ToList();
        }

        public async Task<List<ItemStockAdjustmentDto>> GetItemStockAdjustmentsAsync(string orderNo, string itemCode, string? compItemCode = null)
        {
            if (string.IsNullOrWhiteSpace(orderNo))
                return new List<ItemStockAdjustmentDto>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            var sb = new StringBuilder(@"
                SELECT 
                    soa.EntryID,
                    soa.OrderNo,
                    soa.ItemID,
                    soa.Qty,
                    soa.UserName,
                    soa.DTEntry,
                    ISNULL(soa.Shelf_RefID, 0) AS Shelf_RefID,
                    ISNULL(vrd.LotNo, '') AS LotNo,
                    vrd.NextProcessID,
                    ISNULL(pNext.Description, ISNULL(pCurr.Description, 'Initial Process')) AS StartingProcessName
                FROM StockOrderAdjustments soa
                LEFT JOIN VendRcvdDetail vrd ON soa.VID_RefID = vrd.EntryID
                LEFT JOIN Processes pNext ON vrd.NextProcessID = pNext.ProcessID
                LEFT JOIN Processes pCurr ON vrd.ProcessID = pCurr.ProcessID
                WHERE (soa.OrderNo = @OrderNo OR soa.OrderNo = @CleanOrderNo)");

            var p = new DynamicParameters();
            p.Add("@OrderNo", orderNo.Trim());
            p.Add("@CleanOrderNo", orderNo.Trim().StartsWith("SO-") ? orderNo.Trim().Substring(3) : orderNo.Trim());

            if (!string.IsNullOrWhiteSpace(itemCode))
            {
                sb.Append(@" AND (
                    soa.ItemID = @ItemCode 
                    OR (@CompItemCode IS NOT NULL AND @CompItemCode <> '' AND soa.ItemID = @CompItemCode)
                )");
                p.Add("@ItemCode", itemCode.Trim());
                p.Add("@CompItemCode", compItemCode?.Trim());
            }

            sb.Append(" ORDER BY soa.EntryID DESC");

            return (await db.QueryAsync<ItemStockAdjustmentDto>(sb.ToString(), p)).ToList();
        }
    }
}
