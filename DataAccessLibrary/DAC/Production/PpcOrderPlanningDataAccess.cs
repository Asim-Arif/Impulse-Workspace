using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Production
{
    public class PpcOrderPlanningDataAccess : IPpcOrderPlanningDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<PpcOrderPlanningDataAccess> _logger;

        public PpcOrderPlanningDataAccess(IConfiguration configuration, ILogger<PpcOrderPlanningDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<PpcOrderSummaryListItem>> GetAuthorizedOrdersForPlanningAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        o.OrderNo,
                        o.DT AS OrderDate,
                        o.DeliveryDT AS DeliveryDate,
                        ISNULL(o.CustCode, '') AS CustCode,
                        ISNULL(fc.Name, '') AS CustomerName,
                        ISNULL(o.Country, '') AS Country,
                        COUNT(DISTINCT oi.CompItemCode) AS TotalItems,
                        ISNULL(SUM(oi.Qty), 0) AS TotalOrderQty,
                        CASE WHEN pm.OrderNo IS NOT NULL THEN 1 ELSE 0 END AS IsPlanned,
                        pm.PlannedBy,
                        pm.PlannedAt
                    FROM FCustomerOrders o
                    LEFT JOIN ForeignCustomers fc ON o.CustCode = fc.CustCode AND o.Country = fc.Country
                    LEFT JOIN FOrderItems oi ON o.OrderNo = oi.OrderNo
                    LEFT JOIN PPC_Order_Planning_Master pm ON o.OrderNo = pm.OrderNo
                    WHERE o.Authorized = 1
                    GROUP BY o.OrderNo, o.DT, o.DeliveryDT, o.CustCode, fc.Name, o.Country, pm.OrderNo, pm.PlannedBy, pm.PlannedAt
                    ORDER BY o.DT DESC, o.OrderNo DESC";

                var list = await db.QueryAsync<PpcOrderSummaryListItem>(sql);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching authorized orders for PPC planning");
                throw;
            }
        }

        public async Task<PpcOrderHeaderDto?> GetOrderForPlanningAsync(string orderNo)
        {
            if (string.IsNullOrWhiteSpace(orderNo)) return null;

            try
            {
                using var db = CreateConnection();

                const string headerSql = @"
                    SELECT 
                        o.OrderNo,
                        o.DT AS OrderDate,
                        o.DeliveryDT AS DeliveryDate,
                        ISNULL(o.CustCode, '') AS CustCode,
                        ISNULL(fc.Name, '') AS CustomerName,
                        ISNULL(o.Country, '') AS Country,
                        ISNULL(fc.Curr, '') AS Currency,
                        o.Authorized,
                        CASE WHEN pm.OrderNo IS NOT NULL THEN 1 ELSE 0 END AS IsPlanned,
                        pm.PlannedBy,
                        pm.PlannedAt,
                        pm.Notes
                    FROM FCustomerOrders o
                    LEFT JOIN ForeignCustomers fc ON o.CustCode = fc.CustCode AND o.Country = fc.Country
                    LEFT JOIN PPC_Order_Planning_Master pm ON o.OrderNo = pm.OrderNo
                    WHERE o.OrderNo = @OrderNo";

                var header = await db.QueryFirstOrDefaultAsync<PpcOrderHeaderDto>(headerSql, new { OrderNo = orderNo });
                if (header == null) return null;

                const string itemsSql = @"
                    SELECT 
                        v.ItemCode,
                        v.CompItemID,
                        ISNULL(v.ItemName, v.Description) AS ItemName,
                        v.Qty AS OrderQty,
                        ISNULL(v.InHand, 0) AS AvailableInHandStock,
                        COALESCE(ipg.PG_RefID, i.GroupID, 0) AS GroupID,
                        ISNULL(pg.GroupName, '') AS GroupName
                    FROM VrptOrders_ForProduction v
                    LEFT JOIN Items i ON v.CompItemID = i.ItemID
                    LEFT JOIN ItemProcessGroups ipg ON v.CompItemID = ipg.ItemID
                    LEFT JOIN ProcessGroups pg ON COALESCE(ipg.PG_RefID, i.GroupID) = pg.EntryID
                    WHERE v.OrderNo = @OrderNo
                    ORDER BY v.CompItemID";

                var items = (await db.QueryAsync<PpcOrderItemDto>(itemsSql, new { OrderNo = orderNo })).ToList();
                header.Items = items;

                if (!items.Any()) return header;

                var itemIds = items.Select(i => i.CompItemID).Distinct().ToList();
                var groupIds = items.Where(i => i.GroupID > 0).Select(i => i.GroupID).Distinct().ToList();

                // 1. Fetch saved item planning if previously saved
                const string savedPlanningSql = @"
                    SELECT ItemID, StockQty, TotalPurchaseQty, ProductionQty
                    FROM PPC_Order_Item_Planning
                    WHERE OrderNo = @OrderNo";
                var savedPlanning = (await db.QueryAsync(savedPlanningSql, new { OrderNo = orderNo })).ToList();
                var savedPlanMap = savedPlanning.ToDictionary(x => (string)x.ItemID, x => x);

                // 2. Fetch saved purchases
                const string savedPurchasesSql = @"
                    SELECT 
                        pip.ItemID,
                        pip.ProcessID,
                        p.Description AS ProcessName,
                        pip.VendID,
                        m.VenderName AS MakerName,
                        pip.PurchaseQty,
                        pip.PurchaseRate
                    FROM PPC_Order_Item_Purchases pip
                    INNER JOIN Processes p ON pip.ProcessID = p.ProcessID
                    LEFT JOIN Makers m ON pip.VendID = m.VendID
                    WHERE pip.OrderNo = @OrderNo";
                var savedPurchases = (await db.QueryAsync(savedPurchasesSql, new { OrderNo = orderNo })).ToList();

                // 3. Fetch saved hub schedules
                const string savedSchedulesSql = @"
                    SELECT ItemID, Hub_Name, PlannedQty, StartDate, EndDate
                    FROM PPC_Order_Item_Hub_Schedules
                    WHERE OrderNo = @OrderNo";
                var savedSchedules = (await db.QueryAsync(savedSchedulesSql, new { OrderNo = orderNo })).ToList();

                // 4. Fetch eligible purchase processes for items
                const string purchaseProcSql = @"
                    SELECT DISTINCT ip.ItemID, ip.ProcessID, p.Description AS ProcessName
                    FROM ItemProcesses ip
                    INNER JOIN Processes p ON ip.ProcessID = p.ProcessID
                    INNER JOIN Processes_Purchase pp ON ip.ProcessID = pp.ProcessID
                    WHERE ip.ItemID IN @ItemIds
                    ORDER BY p.Description";
                var purchaseProcRows = (await db.QueryAsync(purchaseProcSql, new { ItemIds = itemIds })).ToList();

                // 5. Fetch distinct hubs for groups
                List<dynamic> groupHubs = new();
                List<HubSupervisorDto> groupSupervisors = new();

                if (groupIds.Any())
                {
                    const string groupHubsSql = @"
                        SELECT DISTINCT pgp.Group_RefID AS GroupID, pgp.Hub_Name
                        FROM ProcessGroupsProcesses pgp
                        WHERE pgp.Group_RefID IN @GroupIds";
                    groupHubs = (await db.QueryAsync(groupHubsSql, new { GroupIds = groupIds })).ToList();

                    const string supsSql = @"
                        SELECT 
                            s.GroupID,
                            s.Hub_Name,
                            s.UserID,
                            s.UserName,
                            ISNULL(u.EmpID, '') AS EmpID,
                            ISNULL(e.Name, u.UserName) AS EmployeeName,
                            ISNULL(e.Designation, '') AS Designation
                        FROM ProcessGroup_Hub_Supervisors s
                        INNER JOIN Users u ON s.UserID = u.UserID
                        LEFT JOIN Employees e ON u.EmpID = e.EmpID
                        WHERE s.GroupID IN @GroupIds";
                    groupSupervisors = (await db.QueryAsync<HubSupervisorDto>(supsSql, new { GroupIds = groupIds })).ToList();
                }

                // Assemble item details
                foreach (var item in items)
                {
                    // Available purchase processes
                    item.AvailablePurchaseProcesses = purchaseProcRows
                        .Where(r => (string)r.ItemID == item.CompItemID)
                        .Select(r => new PurchaseProcessLookupDto
                        {
                            ProcessID = (int)r.ProcessID,
                            ProcessName = (string)r.ProcessName
                        }).ToList();

                    // Restore saved planning values if available
                    if (savedPlanMap.TryGetValue(item.CompItemID, out var plan))
                    {
                        item.StockQty = (int)plan.StockQty;
                    }

                    // Restore saved purchases
                    var itemPurchases = savedPurchases.Where(p => (string)p.ItemID == item.CompItemID).ToList();
                    foreach (var p in itemPurchases)
                    {
                        var alloc = new PpcItemPurchaseAllocationDto
                        {
                            ProcessID = (int)p.ProcessID,
                            ProcessName = (string)p.ProcessName,
                            VendID = (int?)p.VendID,
                            MakerName = (string?)p.MakerName,
                            PurchaseQty = (int)p.PurchaseQty,
                            PurchaseRate = (double)p.PurchaseRate
                        };

                        // Pre-populate available makers for this process
                        alloc.AvailableMakers = await GetMakersForItemAndProcessInternalAsync(db, item.CompItemID, alloc.ProcessID);
                        item.Purchases.Add(alloc);
                    }

                    // Configure hub schedules for this item
                    if (item.GroupID > 0)
                    {
                        var itemHubNames = groupHubs
                            .Where(h => (int)h.GroupID == item.GroupID)
                            .Select(h => (string)h.Hub_Name)
                            .Distinct()
                            .OrderBy(h => h)
                            .ToList();

                        foreach (var hubName in itemHubNames)
                        {
                            var sched = new PpcItemHubScheduleDto
                            {
                                Hub_Name = hubName,
                                PlannedQty = item.ProductionQty,
                                StartDate = DateTime.Today,
                                EndDate = header.DeliveryDate ?? DateTime.Today.AddDays(14)
                            };

                            // Check if previously saved
                            var savedSched = savedSchedules.FirstOrDefault(s => (string)s.ItemID == item.CompItemID && (string)s.Hub_Name == hubName);
                            if (savedSched != null)
                            {
                                sched.PlannedQty = (int)savedSched.PlannedQty;
                                sched.StartDate = (DateTime)savedSched.StartDate;
                                sched.EndDate = (DateTime)savedSched.EndDate;
                            }

                            // Attach supervisors for this group and hub
                            sched.Supervisors = groupSupervisors
                                .Where(s => s.GroupID == item.GroupID && s.Hub_Name == hubName)
                                .ToList();

                            item.HubSchedules.Add(sched);
                        }
                    }
                }

                return header;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching order #{OrderNo} for PPC planning", orderNo);
                throw;
            }
        }

        public async Task<List<PurchaseProcessLookupDto>> GetPurchaseProcessesForItemAsync(string itemId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT DISTINCT ip.ProcessID, p.Description AS ProcessName
                    FROM ItemProcesses ip
                    INNER JOIN Processes p ON ip.ProcessID = p.ProcessID
                    INNER JOIN Processes_Purchase pp ON ip.ProcessID = pp.ProcessID
                    WHERE ip.ItemID = @ItemId
                    ORDER BY p.Description";

                var list = await db.QueryAsync<PurchaseProcessLookupDto>(sql, new { ItemId = itemId });
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching purchase processes for item {ItemId}", itemId);
                return new List<PurchaseProcessLookupDto>();
            }
        }

        public async Task<List<MakerLookupDto>> GetMakersForItemAndProcessAsync(string itemId, int processId)
        {
            try
            {
                using var db = CreateConnection();
                return await GetMakersForItemAndProcessInternalAsync(db, itemId, processId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching makers for item {ItemId} and process {ProcessId}", itemId, processId);
                return new List<MakerLookupDto>();
            }
        }

        private async Task<List<MakerLookupDto>> GetMakersForItemAndProcessInternalAsync(IDbConnection db, string itemId, int processId)
        {
            const string sql = @"
                SELECT DISTINCT m.VendID, m.VenderName AS MakerName, ISNULL(vai.Rate, 0) AS Rate
                FROM Makers m
                INNER JOIN VendAssItems vai ON m.VendID = vai.VendID
                WHERE vai.ItemID = @ItemId AND vai.ProcessID = @ProcessId AND m.Active = 1
                ORDER BY m.VenderName";

            var list = await db.QueryAsync<MakerLookupDto>(sql, new { ItemId = itemId, ProcessId = processId });
            return list.ToList();
        }

        public async Task<bool> SavePpcPlanAsync(PpcPlanSaveRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.OrderNo)) return false;

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var trans = conn.BeginTransaction();

            try
            {
                var orderNo = request.OrderNo.Trim();
                var user = request.PlannedBy ?? "System";

                // 1. Delete previous records for this order
                await conn.ExecuteAsync("DELETE FROM PPC_Order_Item_Hub_Schedules WHERE OrderNo = @OrderNo", new { OrderNo = orderNo }, trans);
                await conn.ExecuteAsync("DELETE FROM PPC_Order_Item_Purchases WHERE OrderNo = @OrderNo", new { OrderNo = orderNo }, trans);
                await conn.ExecuteAsync("DELETE FROM PPC_Order_Item_Planning WHERE OrderNo = @OrderNo", new { OrderNo = orderNo }, trans);
                await conn.ExecuteAsync("DELETE FROM PPC_Order_Planning_Master WHERE OrderNo = @OrderNo", new { OrderNo = orderNo }, trans);

                // 2. Insert Master record
                const string masterSql = @"
                    INSERT INTO PPC_Order_Planning_Master (OrderNo, PlannedBy, PlannedAt, Status, Notes)
                    VALUES (@OrderNo, @PlannedBy, GETUTCDATE(), 'Planned', @Notes)";

                await conn.ExecuteAsync(masterSql, new
                {
                    OrderNo = orderNo,
                    PlannedBy = user,
                    Notes = request.Notes
                }, trans);

                // 3. Insert Item Planning & Purchases & Hub Schedules
                const string itemPlanSql = @"
                    INSERT INTO PPC_Order_Item_Planning 
                        (OrderNo, ItemID, OrderQty, StockQty, TotalPurchaseQty, ProductionQty, Remarks, CreatedAt, CreatedBy)
                    VALUES 
                        (@OrderNo, @ItemID, @OrderQty, @StockQty, @TotalPurchaseQty, @ProductionQty, @Remarks, GETUTCDATE(), @CreatedBy)";

                const string purchaseSql = @"
                    INSERT INTO PPC_Order_Item_Purchases 
                        (OrderNo, ItemID, ProcessID, VendID, PurchaseQty, PurchaseRate, CreatedAt, CreatedBy)
                    VALUES 
                        (@OrderNo, @ItemID, @ProcessID, @VendID, @PurchaseQty, @PurchaseRate, GETUTCDATE(), @CreatedBy)";

                const string schedSql = @"
                    INSERT INTO PPC_Order_Item_Hub_Schedules 
                        (OrderNo, ItemID, Hub_Name, PlannedQty, StartDate, EndDate, CreatedAt, CreatedBy)
                    VALUES 
                        (@OrderNo, @ItemID, @Hub_Name, @PlannedQty, @StartDate, @EndDate, GETUTCDATE(), @CreatedBy)";

                foreach (var item in request.Items)
                {
                    // Item summary row
                    await conn.ExecuteAsync(itemPlanSql, new
                    {
                        OrderNo = orderNo,
                        ItemID = item.CompItemID,
                        OrderQty = item.OrderQty,
                        StockQty = item.StockQty,
                        TotalPurchaseQty = item.TotalPurchaseQty,
                        ProductionQty = item.ProductionQty,
                        Remarks = "",
                        CreatedBy = user
                    }, trans);

                    // Multi-purchase split rows
                    foreach (var p in item.Purchases.Where(x => x.PurchaseQty > 0 && x.ProcessID > 0))
                    {
                        await conn.ExecuteAsync(purchaseSql, new
                        {
                            OrderNo = orderNo,
                            ItemID = item.CompItemID,
                            ProcessID = p.ProcessID,
                            VendID = p.VendID,
                            PurchaseQty = p.PurchaseQty,
                            PurchaseRate = p.PurchaseRate,
                            CreatedBy = user
                        }, trans);
                    }

                    // Hub schedule rows
                    foreach (var hub in item.HubSchedules)
                    {
                        await conn.ExecuteAsync(schedSql, new
                        {
                            OrderNo = orderNo,
                            ItemID = item.CompItemID,
                            Hub_Name = hub.Hub_Name,
                            PlannedQty = hub.PlannedQty,
                            StartDate = hub.StartDate,
                            EndDate = hub.EndDate,
                            CreatedBy = user
                        }, trans);
                    }
                }

                trans.Commit();
                _logger.LogInformation("PPC Plan successfully saved for Order #{OrderNo} with {ItemCount} items.", orderNo, request.Items.Count);
                return true;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                _logger.LogError(ex, "Failed to save PPC Plan for Order #{OrderNo}", request.OrderNo);
                throw;
            }
        }
    }
}
