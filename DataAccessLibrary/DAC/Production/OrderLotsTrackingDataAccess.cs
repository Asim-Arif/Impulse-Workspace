using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Production
{
    public class OrderLotsTrackingDataAccess : IOrderLotsTrackingDataAccess
    {
        private readonly IConfiguration _config;
        private readonly ILogger<OrderLotsTrackingDataAccess> _logger;

        public OrderLotsTrackingDataAccess(IConfiguration config, ILogger<OrderLotsTrackingDataAccess> logger)
        {
            _config = config;
            _logger = logger;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("SMBI_AWM")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<List<TrackingCustomerLookupItem>> GetCustomersLookupAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT DISTINCT 
                    co.CustCode, 
                    ISNULL(c.Name, co.CustCode) AS CustName
                FROM FCustomerOrders co WITH (NOLOCK)
                LEFT JOIN ForeignCustomers c WITH (NOLOCK) ON co.CustCode = c.CustCode
                WHERE ISNULL(co.CustCode, '') <> ''
                ORDER BY co.CustCode ASC";

            return (await db.QueryAsync<TrackingCustomerLookupItem>(sql)).ToList();
        }

        public async Task<List<TrackingOrderLookupItem>> GetOrdersLookupAsync(string? custCode = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT TOP 400
                    co.OrderNo,
                    ISNULL(co.InternalRefNo, '') AS InternalRefNo,
                    co.CustCode,
                    co.DT
                FROM FCustomerOrders co WITH (NOLOCK)
                WHERE (@CustCode IS NULL OR @CustCode = '' OR co.CustCode = @CustCode)
                ORDER BY co.DT DESC, co.InternalRefNo DESC, co.OrderNo DESC";

            return (await db.QueryAsync<TrackingOrderLookupItem>(sql, new { CustCode = custCode })).ToList();
        }

        public async Task<List<string>> GetHubNamesLookupAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT DISTINCT Hub_Name 
                FROM Hub_Names WITH (NOLOCK)
                WHERE ISNULL(Hub_Name, '') <> ''
                UNION
                SELECT DISTINCT Hub_Name 
                FROM ProcessGroupsProcesses WITH (NOLOCK)
                WHERE ISNULL(Hub_Name, '') <> ''
                ORDER BY Hub_Name ASC";

            return (await db.QueryAsync<string>(sql)).ToList();
        }

        public async Task<OrderLotsTrackingDashboardDto> GetOrderLotsTrackingAsync(OrderLotsTrackingFilter filter)
        {
            var result = new OrderLotsTrackingDashboardDto();

            using IDbConnection db = new SqlConnection(ConnectionString);
            var p = new DynamicParameters();

            // Date filtering parameters
            DateTime? startDate = null;
            DateTime? endDate = null;
            DateTime now = DateTime.Today;

            switch (filter.DateRangeType)
            {
                case 1: // Today
                    startDate = now;
                    endDate = now.AddDays(1).AddTicks(-1);
                    break;
                case 2: // Last 15 Days
                    startDate = now.AddDays(-15);
                    endDate = now.AddDays(1).AddTicks(-1);
                    break;
                case 3: // Last 30 Days
                    startDate = now.AddDays(-30);
                    endDate = now.AddDays(1).AddTicks(-1);
                    break;
                case 4: // Last 60 Days
                    startDate = now.AddDays(-60);
                    endDate = now.AddDays(1).AddTicks(-1);
                    break;
                case 5: // Last 90 Days
                    startDate = now.AddDays(-90);
                    endDate = now.AddDays(1).AddTicks(-1);
                    break;
                case 6: // Custom
                    if (filter.DtFrom.HasValue) startDate = filter.DtFrom.Value.Date;
                    if (filter.DtTo.HasValue) endDate = filter.DtTo.Value.Date.AddDays(1).AddTicks(-1);
                    break;
            }

            p.Add("@StartDate", startDate);
            p.Add("@EndDate", endDate);

            // Customer
            if (!string.IsNullOrWhiteSpace(filter.CustCode))
            {
                p.Add("@CustCode", filter.CustCode.Trim());
            }

            // OrderNo / InternalRefNo
            if (!string.IsNullOrWhiteSpace(filter.OrderNo))
            {
                string orderNo = filter.OrderNo.Trim();
                string cleanOrderNo = orderNo.StartsWith("SO-", StringComparison.OrdinalIgnoreCase)
                    ? orderNo.Substring(3)
                    : orderNo;

                p.Add("@OrderNo", orderNo);
                p.Add("@CleanOrderNo", cleanOrderNo);
            }
            if (!string.IsNullOrWhiteSpace(filter.InternalRefNo))
            {
                p.Add("@InternalRefNo", filter.InternalRefNo.Trim());
            }

            // Hub
            if (!string.IsNullOrWhiteSpace(filter.HubName))
            {
                p.Add("@HubName", filter.HubName.Trim());
            }

            // Completed lots union clause
            string completedUnion = filter.IncludeCompleted ? @"
                UNION ALL

                -- 3. Lots finished in production
                SELECT 
                    vrd.OrderNo,
                    vrd.ItemCode,
                    vrd.LotNo,
                    vrd.ProcessID,
                    vr.VendID,
                    vr.DT AS LastActivityDT,
                    ISNULL(vrd.ReWorkLot, 0) AS ReWorkLot,
                    vrd.RcvdQty AS Qty,
                    ip.SNO,
                    'Completed' AS LotState,
                    1 AS IsCompleted
                FROM VendRcvdDetail vrd WITH (NOLOCK)
                INNER JOIN VendReceived vr WITH (NOLOCK) ON vrd.RefID = vr.EntryID
                LEFT JOIN ItemProcesses ip WITH (NOLOCK) ON vrd.ItemCode = ip.ItemID AND vrd.ProcessID = ip.ProcessID
                WHERE vrd.LotNo <> '0'
                  AND ((vrd.NextProcessID IS NULL OR vrd.NextProcessID = 0)
                       OR EXISTS (SELECT 1 FROM Lots_Closed lc WITH (NOLOCK) WHERE lc.LotNo = vrd.LotNo))" : "";

            var queryBuilder = new StringBuilder($@"
                ;WITH ProcessStepCounts AS (
                    SELECT Group_RefID, COUNT(1) AS TotalSteps
                    FROM ProcessGroupsProcesses WITH (NOLOCK)
                    GROUP BY Group_RefID
                ),
                OrderLots AS (
                    -- 1. Lots issued to maker and currently in process
                    SELECT 
                        vid.OrderNo,
                        vid.ItemCode,
                        vid.LotNo,
                        vid.RcvProcessID AS ProcessID,
                        vi.VendID,
                        vi.DT AS LastActivityDT,
                        ISNULL(vid.ReWorkLot, 0) AS ReWorkLot,
                        (vid.IssQty - vid.RcvdQty) AS Qty,
                        ip.SNO,
                        'Issued to Maker' AS LotState,
                        0 AS IsCompleted
                    FROM VendIssdDetail vid WITH (NOLOCK)
                    INNER JOIN VendIssued vi WITH (NOLOCK) ON vid.RefID = vi.EntryID
                    LEFT JOIN ItemProcesses ip WITH (NOLOCK) ON vid.ItemCode = ip.ItemID AND vid.RcvProcessID = ip.ProcessID
                    WHERE vid.LotNo <> '0'
                      AND vid.IssQty > vid.RcvdQty
                      AND NOT EXISTS (
                          SELECT 1 FROM VendRcvdDetail vrd WITH (NOLOCK) 
                          WHERE vrd.Issue_RefID = vid.EntryID AND vrd.OrderNo = vid.OrderNo
                      )
                      AND NOT EXISTS (
                          SELECT 1 FROM Lots_Closed lc WITH (NOLOCK) 
                          WHERE lc.LotNo = vid.LotNo
                      )

                    UNION ALL

                    -- 2. Lots received from maker but waiting for next process
                    SELECT 
                        vrd.OrderNo,
                        vrd.ItemCode,
                        vrd.LotNo,
                        vrd.ProcessID,
                        vr.VendID,
                        vr.DT AS LastActivityDT,
                        ISNULL(vrd.ReWorkLot, 0) AS ReWorkLot,
                        (vrd.RcvdQty - vrd.IssQty - ISNULL(vrd.Wastage, 0) - ISNULL(vrd.ReWorkQty, 0)) AS Qty,
                        ip.SNO,
                        'In Hub (Received)' AS LotState,
                        0 AS IsCompleted
                    FROM VendRcvdDetail vrd WITH (NOLOCK)
                    INNER JOIN VendReceived vr WITH (NOLOCK) ON vrd.RefID = vr.EntryID
                    LEFT JOIN ItemProcesses ip WITH (NOLOCK) ON vrd.ItemCode = ip.ItemID AND vrd.ProcessID = ip.ProcessID
                    WHERE vrd.LotNo <> '0'
                      AND (vrd.RcvdQty - vrd.IssQty - ISNULL(vrd.Wastage, 0) - ISNULL(vrd.ReWorkQty, 0)) > 0
                      AND (vrd.NextProcessID IS NOT NULL AND vrd.NextProcessID <> 0)
                      AND ISNULL(vrd.Opening_RefID, 0) = 0
                      AND NOT EXISTS (
                          SELECT 1 FROM VendIssdDetail vid WITH (NOLOCK) 
                          WHERE vid.Rcvd_RefID = vrd.EntryID AND vid.OrderNo = vrd.OrderNo
                      )
                      AND NOT EXISTS (
                          SELECT 1 FROM Lots_Closed lc WITH (NOLOCK) 
                          WHERE lc.LotNo = vrd.LotNo
                      )

                    {completedUnion}
                ),
                LotSteps AS (
                    SELECT 
                        ol.OrderNo,
                        ol.ItemCode,
                        ol.LotNo,
                        ol.ProcessID,
                        ol.VendID,
                        ol.LastActivityDT,
                        ol.ReWorkLot,
                        ol.Qty,
                        ol.SNO,
                        ol.LotState,
                        ol.IsCompleted,
                        MAX(ol.SNO) OVER (PARTITION BY ol.LotNo) AS MaxSno
                    FROM OrderLots ol
                )
                SELECT 
                    ls.LotNo,
                    ls.ItemCode,
                    ISNULL(i.ItemName, ls.ItemCode) AS ItemName,
                    ISNULL(i.ItemSize, '') AS ItemSize,
                    ls.OrderNo,
                    ISNULL(co.InternalRefNo, '') AS InternalRefNo,
                    co.CustCode,
                    ISNULL(c.Name, co.CustCode) AS CustomerName,
                    ls.ProcessID,
                    ISNULL(p.Description, '') AS ProcessName,
                    ISNULL(currPgp.Hub_Name, 'Unassigned Hub') AS Hub_Name,
                    ISNULL(sup.Supervisors, '') AS Supervisors,
                    ISNULL(ls.Qty, 0) AS Qty,
                    COALESCE(currPgp.SeqNo, CAST(ls.SNO AS INT), 1) AS CurrentSeqNo,
                    COALESCE(psc.TotalSteps, CAST(ls.MaxSno AS INT), 1) AS TotalSeqNo,
                    ls.LastActivityDT,
                    ISNULL(m.VenderName, 'In-House Production') AS MakerName,
                    ls.LotState,
                    ls.IsCompleted,
                    ls.ReWorkLot AS IsReWork,
                    COALESCE(ppcSched.EndDate, foi.DeliveryDT, co.DeliveryDT) AS TargetDate,
                    ppcSched.StartDate AS PlannedStartDate,
                    co.DeliveryDT AS OrderDeliveryDate
                FROM LotSteps ls
                INNER JOIN FCustomerOrders co WITH (NOLOCK) ON ls.OrderNo = co.OrderNo
                LEFT JOIN ForeignCustomers c WITH (NOLOCK) ON co.CustCode = c.CustCode
                LEFT JOIN Items i WITH (NOLOCK) ON ls.ItemCode = i.ItemID
                LEFT JOIN Processes p WITH (NOLOCK) ON ls.ProcessID = p.ProcessID
                LEFT JOIN Makers m WITH (NOLOCK) ON ls.VendID = m.VendID
                LEFT JOIN ItemProcessGroups ipg WITH (NOLOCK) ON ls.ItemCode = ipg.ItemID
                LEFT JOIN ProcessGroupsProcesses currPgp WITH (NOLOCK) ON ipg.PG_RefID = currPgp.Group_RefID AND ls.ProcessID = currPgp.Process_RefID
                LEFT JOIN ProcessStepCounts psc ON ipg.PG_RefID = psc.Group_RefID
                LEFT JOIN PPC_Order_Item_Hub_Schedules ppcSched WITH (NOLOCK) 
                    ON ppcSched.OrderNo = ls.OrderNo 
                   AND ppcSched.ItemID = ls.ItemCode 
                   AND ppcSched.Hub_Name = currPgp.Hub_Name
                LEFT JOIN FOrderItems foi WITH (NOLOCK) 
                    ON foi.OrderNo = ls.OrderNo 
                   AND (foi.ItemCode = ls.ItemCode OR foi.CompItemCode = ls.ItemCode)
                OUTER APPLY (
                    SELECT STUFF((
                        SELECT ', ' + pghs.UserName
                        FROM ProcessGroup_Hub_Supervisors pghs WITH (NOLOCK)
                        WHERE pghs.GroupID = ipg.PG_RefID AND pghs.Hub_Name = currPgp.Hub_Name
                        FOR XML PATH(''), TYPE
                    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS Supervisors
                ) sup
                WHERE 1 = 1");

            // Filter Conditions
            if (!string.IsNullOrWhiteSpace(filter.CustCode))
            {
                queryBuilder.Append(" AND co.CustCode = @CustCode");
            }

            if (!string.IsNullOrWhiteSpace(filter.InternalRefNo))
            {
                queryBuilder.Append(" AND (co.InternalRefNo = @InternalRefNo OR ls.OrderNo = @InternalRefNo)");
            }
            else if (!string.IsNullOrWhiteSpace(filter.OrderNo))
            {
                queryBuilder.Append(" AND (ls.OrderNo = @OrderNo OR ls.OrderNo = @CleanOrderNo OR co.InternalRefNo = @OrderNo)");
            }

            if (!string.IsNullOrWhiteSpace(filter.HubName))
            {
                queryBuilder.Append(" AND currPgp.Hub_Name = @HubName");
            }

            if (startDate.HasValue && endDate.HasValue)
            {
                if (filter.DateFilterMode == 1) // Target Date
                {
                    queryBuilder.Append(" AND COALESCE(ppcSched.EndDate, foi.DeliveryDT, co.DeliveryDT) BETWEEN @StartDate AND @EndDate");
                }
                else // Order Date
                {
                    queryBuilder.Append(" AND co.DT BETWEEN @StartDate AND @EndDate");
                }
            }

            queryBuilder.Append(" ORDER BY ls.OrderNo DESC, ls.LotNo ASC");

            var lots = (await db.QueryAsync<OrderLotTrackingItemDto>(queryBuilder.ToString(), p)).ToList();
            result.Lots = lots;

            // Summary Metrics
            result.TotalLots = lots.Count;
            result.TotalQty = lots.Sum(x => x.Qty);
            result.ActiveWipLots = lots.Count(x => !x.IsCompleted);
            result.ActiveWipQty = lots.Where(x => !x.IsCompleted).Sum(x => x.Qty);
            result.CompletedLots = lots.Count(x => x.IsCompleted);

            result.OverdueLots = lots.Count(x => x.TargetStatus == "Overdue");
            result.DueSoonLots = lots.Count(x => x.TargetStatus == "Due Soon" || x.TargetStatus == "Due Today");
            result.OnTrackLots = lots.Count(x => x.TargetStatus == "On Track");

            // Hub Summaries Breakdown
            result.HubSummaries = lots
                .GroupBy(x => string.IsNullOrWhiteSpace(x.Hub_Name) ? "Unassigned" : x.Hub_Name, StringComparer.OrdinalIgnoreCase)
                .Select(g => new OrderLotsHubSummaryDto
                {
                    Hub_Name = g.Key,
                    LotCount = g.Count(),
                    TotalQty = g.Sum(x => x.Qty),
                    OverdueCount = g.Count(x => x.TargetStatus == "Overdue"),
                    Supervisors = g.Select(x => x.Supervisors).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? string.Empty
                })
                .OrderBy(x => x.Hub_Name)
                .ToList();

            return result;
        }
    }
}
