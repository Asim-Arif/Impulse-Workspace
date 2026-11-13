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

        private static bool _schemaEnsured = false;

        private async Task EnsureExceptionSchemaAsync(IDbConnection db)
        {
            if (_schemaEnsured) return;
            try
            {
                const string ddl = @"
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Planning') AND name = 'RejectedStockQty')
                        ALTER TABLE PPC_Order_Item_Planning ADD RejectedStockQty INT NOT NULL CONSTRAINT DF_PPC_ItemPlan_RejectedStockQty DEFAULT 0;
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Planning') AND name = 'RejectedSFQty')
                        ALTER TABLE PPC_Order_Item_Planning ADD RejectedSFQty INT NOT NULL CONSTRAINT DF_PPC_ItemPlan_RejectedSFQty DEFAULT 0;
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'Status')
                        ALTER TABLE PPC_Order_Item_Purchases ADD Status VARCHAR(30) NOT NULL CONSTRAINT DF_PPC_ItemPurch_Status DEFAULT 'Active';
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'OriginalPlannedQty')
                        ALTER TABLE PPC_Order_Item_Purchases ADD OriginalPlannedQty INT NULL;
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'ShortClosedQty')
                        ALTER TABLE PPC_Order_Item_Purchases ADD ShortClosedQty INT NOT NULL CONSTRAINT DF_PPC_ItemPurch_ShortClosedQty DEFAULT 0;
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'ExceptionNotes')
                        ALTER TABLE PPC_Order_Item_Purchases ADD ExceptionNotes NVARCHAR(500) NULL;
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'UpdatedAt')
                        ALTER TABLE PPC_Order_Item_Purchases ADD UpdatedAt DATETIME2 NULL;
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'UpdatedBy')
                        ALTER TABLE PPC_Order_Item_Purchases ADD UpdatedBy NVARCHAR(100) NULL;
                    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'PPC_Order_Planning_Exceptions')
                    BEGIN
                        CREATE TABLE dbo.PPC_Order_Planning_Exceptions (
                            EntryID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
                            OrderNo VARCHAR(100) NOT NULL,
                            ItemID VARCHAR(50) NOT NULL,
                            ExceptionType VARCHAR(50) NOT NULL,
                            AffectedQty INT NOT NULL DEFAULT 0,
                            SourceRef VARCHAR(100) NULL,
                            Reason NVARCHAR(500) NOT NULL,
                            CreatedBy NVARCHAR(100) NOT NULL,
                            CreatedAt DATETIME2 NOT NULL DEFAULT GETUTCDATE()
                        );
                        CREATE INDEX IX_PPC_Exceptions_Order_Item ON dbo.PPC_Order_Planning_Exceptions(OrderNo, ItemID, CreatedAt DESC);
                    END";
                await db.ExecuteAsync(ddl);
                _schemaEnsured = true;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "EnsureExceptionSchemaAsync warning in PpcOrderPlanningDataAccess");
            }
        }

        public async Task<List<PpcOrderSummaryListItem>> GetAuthorizedOrdersForPlanningAsync()
        {
            try
            {
                using var db = CreateConnection();
                await EnsureExceptionSchemaAsync(db);
                const string sql = @"
                    SELECT 
                        o.OrderNo,
                        o.DT AS OrderDate,
                        o.DeliveryDT AS DeliveryDate,
                        ISNULL(o.CustCode, '') AS CustCode,
                        ISNULL(fc.Name, '') AS CustomerName,
                        ISNULL(o.Country, '') AS Country,
                        ISNULL(v_counts.TotalItems, ISNULL(oi_counts.TotalItems, 0)) AS TotalItems,
                        ISNULL(pip_counts.PlannedItems, 0) AS PlannedItems,
                        ISNULL(oi_counts.TotalOrderQty, 0) AS TotalOrderQty,
                        CASE 
                            WHEN pm.Status = 'Planned' 
                             AND ISNULL(oi_counts.TotalOrderQty, 0) <= ISNULL(pip_counts.TotalPlannedQty, 0)
                             AND ISNULL(pip_counts.PlannedItems, 0) >= ISNULL(v_counts.TotalItems, ISNULL(oi_counts.TotalItems, 0))
                            THEN 1 
                            ELSE 0 
                        END AS IsPlanned,
                        pm.PlannedBy,
                        pm.PlannedAt
                    FROM FCustomerOrders o
                    LEFT JOIN ForeignCustomers fc ON o.CustCode = fc.CustCode AND o.Country = fc.Country
                    LEFT JOIN PPC_Order_Planning_Master pm ON o.OrderNo = pm.OrderNo
                    OUTER APPLY (
                        SELECT COUNT(DISTINCT CompItemID) AS TotalItems
                        FROM VrptOrders_ForProduction v
                        WHERE v.OrderNo = o.OrderNo
                    ) v_counts
                    OUTER APPLY (
                        SELECT COUNT(DISTINCT ItemID) AS PlannedItems, SUM(OrderQty) AS TotalPlannedQty
                        FROM PPC_Order_Item_Planning pip
                        WHERE pip.OrderNo = o.OrderNo
                    ) pip_counts
                    OUTER APPLY (
                        SELECT COUNT(DISTINCT CompItemCode) AS TotalItems, SUM(Qty) AS TotalOrderQty
                        FROM FOrderItems oi
                        WHERE oi.OrderNo = o.OrderNo
                    ) oi_counts
                    WHERE o.Authorized = 1 OR EXISTS (SELECT 1 FROM FOrderItems foi WHERE foi.OrderNo = o.OrderNo AND ISNULL(foi.Authorized, 0) = 1)
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
                await EnsureExceptionSchemaAsync(db);

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
                        CASE WHEN pm.Status = 'Planned' THEN 1 ELSE 0 END AS IsPlanned,
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
                        ISNULL(foi.Authorized, 0) AS Authorized,
                        CASE 
                            WHEN v.SetItem = 1 THEN ISNULL(foi.AuthorizedQty, 0) * ISNULL(inset.Qty, 1)
                            ELSE ISNULL(foi.AuthorizedQty, 0)
                        END AS AuthorizedQty,
                        ISNULL(v.InHand, 0) AS AvailableInHandStock,
                        COALESCE(ipg.PG_RefID, i.GroupID, 0) AS GroupID,
                        ISNULL(pg.GroupName, '') AS GroupName
                    FROM VrptOrders_ForProduction v
                    LEFT JOIN (
                        SELECT OrderNo, CompItemCode, 
                               MIN(CAST(ISNULL(Authorized, 0) AS INT)) AS Authorized, 
                               SUM(ISNULL(AuthorizedQty, 0)) AS AuthorizedQty
                        FROM FOrderItems
                        GROUP BY OrderNo, CompItemCode
                    ) foi ON v.OrderNo = foi.OrderNo AND COALESCE(v.OrderItemID, v.CompItemID) = foi.CompItemCode
                    LEFT JOIN ItemsSets inset ON v.OrderItemID = inset.ItemID AND v.CompItemID = inset.Set_ItemID
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
                    SELECT ItemID, OrderQty, StockQty, ISNULL(SFStockQty, 0) AS SFStockQty,
                           ISNULL(RejectedStockQty, 0) AS RejectedStockQty, ISNULL(RejectedSFQty, 0) AS RejectedSFQty,
                           TotalPurchaseQty, ProductionQty
                    FROM PPC_Order_Item_Planning
                    WHERE OrderNo = @OrderNo";
                var savedPlanning = (await db.QueryAsync(savedPlanningSql, new { OrderNo = orderNo })).ToList();
                var savedPlanMap = savedPlanning.ToDictionary(x => (string)x.ItemID, x => x);

                // 1a. Fetch live finish stock already adjusted/withdrawn from store for this order
                const string adjustedStockSql = @"
                    SELECT ItemID, ISNULL(SUM(Qty), 0) AS AlreadyAdjustedQty
                    FROM StockOrderAdjustments
                    WHERE OrderNo = @OrderNo
                    GROUP BY ItemID";
                var adjustedStockRows = (await db.QueryAsync(adjustedStockSql, new { OrderNo = orderNo })).ToList();
                var adjustedStockMap = adjustedStockRows.ToDictionary(x => (string)x.ItemID, x => (int)x.AlreadyAdjustedQty);

                // 1b. Fetch available semi-finished stock lots for items
                const string sfLotsSql = @"
                    SELECT 
                        SOO.EntryID,
                        SOO.ItemID,
                        SOO.ProcessID,
                        ISNULL(SOO.Description, '') AS ProcessName,
                        ISNULL(SOO.LotNo_Manual, '') AS LotNo,
                        ISNULL(SOO.BatchNo_Manual, '') AS BatchNo,
                        ISNULL(SOO.Shelf_RefID, 0) AS Shelf_RefID,
                        (ISNULL(SOO.StoreName, '') + CASE WHEN ISNULL(SOO.RackNo, '') <> '' THEN ' (Rack ' + SOO.RackNo + ', Shelf ' + ISNULL(SOO.ShelfNo, '') + ')' ELSE '' END) AS StoreLocation,
                        CAST((SOO.Qty - ISNULL(SOO.IssdQty, 0)) AS INT) AS AvailableQty
                    FROM VStockOrderOpening SOO
                    WHERE SOO.ItemID IN @ItemIds AND (SOO.Qty - ISNULL(SOO.IssdQty, 0)) > 0
                    ORDER BY SOO.EntryID DESC";
                var availableSfLots = (await db.QueryAsync<PpcItemSFLotDto>(sfLotsSql, new { ItemIds = itemIds })).ToList();

                // 1c. Fetch saved semi-finished stock allocations for this order with live issued qty
                const string savedSfAllocSql = @"
                    SELECT 
                        A.EntryID,
                        A.OrderNo,
                        A.ItemID,
                        A.SOO_RefID,
                        A.ProcessID,
                        ISNULL(A.ProcessName, '') AS ProcessName,
                        ISNULL(A.LotNo, '') AS LotNo,
                        ISNULL(A.BatchNo, '') AS BatchNo,
                        ISNULL(A.Shelf_RefID, 0) AS Shelf_RefID,
                        ISNULL(A.StoreLocation, '') AS StoreLocation,
                        A.AllocatedQty,
                        ISNULL((
                            SELECT SUM(SOOI.QtyIssd)
                            FROM StockOrderOpening_Issuance SOOI
                            INNER JOIN VendIssdDetail VID ON SOOI.VID_RefID = VID.EntryID
                            WHERE VID.OrderNo = A.OrderNo AND SOOI.SOO_RefID = A.SOO_RefID
                        ), 0) AS AlreadyIssuedQty
                    FROM PPC_Order_Item_SF_Allocations A
                    WHERE A.OrderNo = @OrderNo";
                var savedSfAllocations = (await db.QueryAsync<dynamic>(savedSfAllocSql, new { OrderNo = orderNo })).ToList();

                // 2. Fetch saved purchases with live PO execution status
                const string savedPurchasesSql = @"
                    SELECT 
                        pip.EntryID,
                        pip.ItemID,
                        pip.ProcessID,
                        p.Description AS ProcessName,
                        pip.VendID,
                        m.VenderName AS MakerName,
                        pip.PurchaseQty,
                        pip.PurchaseRate,
                        ISNULL(pip.ExcessQtyPercentage, 50) AS ExcessQtyPercentage,
                        COALESCE(pip.VendIssued_RefID, po_exec.VendIssued_RefID) AS VendIssued_RefID,
                        COALESCE(pip.MasterPONo, po_exec.MasterPONo) AS MasterPONo,
                        pip.IssuedAt,
                        pip.IssuedBy,
                        ISNULL(pip.Status, 'Active') AS Status,
                        pip.OriginalPlannedQty,
                        ISNULL(pip.ShortClosedQty, 0) AS ShortClosedQty,
                        pip.ExceptionNotes,
                        po_exec.DetailEntryID AS VendIssdDetail_EntryID,
                        ISNULL(po_exec.IssQty, 0) AS IssuedQty,
                        ISNULL(po_exec.RcvdQty, 0) AS ReceivedQty
                    FROM PPC_Order_Item_Purchases pip
                    INNER JOIN Processes p ON pip.ProcessID = p.ProcessID
                    LEFT JOIN Makers m ON pip.VendID = m.VendID
                    OUTER APPLY (
                        SELECT TOP 1 
                            vid.RefID AS VendIssued_RefID,
                            vi.MasterPONo,
                            vid.EntryID AS DetailEntryID,
                            CAST(ISNULL(vid.IssQty, 0) AS INT) AS IssQty,
                            CAST(ISNULL((SELECT SUM(vrd.RcvdQty) FROM VendRcvdDetail vrd WHERE vrd.Issue_RefID = vid.EntryID), ISNULL(vid.RcvdQty, 0)) AS INT) AS RcvdQty
                        FROM VendIssdDetail vid
                        INNER JOIN VendIssued vi ON vid.RefID = vi.EntryID
                        WHERE vid.OrderNo = pip.OrderNo
                          AND vid.ItemCode = pip.ItemID
                          AND vid.RcvProcessID = pip.ProcessID
                          AND (
                              (pip.VendIssued_RefID IS NOT NULL AND pip.VendIssued_RefID > 0 AND vid.RefID = pip.VendIssued_RefID)
                              OR (pip.VendIssued_RefID IS NULL AND vi.MasterPONo IS NOT NULL AND vi.MasterPONo <> '' AND vid.LotNo = '0')
                          )
                        ORDER BY vid.EntryID DESC
                    ) po_exec
                    WHERE pip.OrderNo = @OrderNo
                      AND ISNULL(pip.Status, 'Active') <> 'Cancelled'";
                var savedPurchases = (await db.QueryAsync(savedPurchasesSql, new { OrderNo = orderNo })).ToList();

                // 2b. Fetch exception audit logs for this order
                const string exceptionsSql = @"
                    SELECT EntryID, OrderNo, ItemID, ExceptionType, AffectedQty, SourceRef, Reason, CreatedBy, CreatedAt
                    FROM PPC_Order_Planning_Exceptions
                    WHERE OrderNo = @OrderNo
                    ORDER BY CreatedAt DESC";
                header.ExceptionLogs = (await db.QueryAsync<PpcPlanningExceptionLogDto>(exceptionsSql, new { OrderNo = orderNo })).ToList();

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

                    // Available SF lots
                    item.AvailableSFLots = availableSfLots
                        .Where(l => l.ItemID == item.CompItemID)
                        .ToList();

                    // Live already-adjusted finish stock and exception logs
                    item.AlreadyAdjustedStockQty = adjustedStockMap.TryGetValue(item.CompItemID, out int adjQty) ? adjQty : 0;
                    item.ItemExceptions = header.ExceptionLogs.Where(e => e.ItemID == item.CompItemID).ToList();

                    // Restore saved planning values if available (only for authorized items)
                    if (item.Authorized && savedPlanMap.TryGetValue(item.CompItemID, out var plan))
                    {
                        var planOrderQty = (int)plan.OrderQty;
                        item.PreviouslyPlannedOrderQty = planOrderQty;
                        item.StockQty = (int)plan.StockQty;
                        item.RejectedStockQty = (int)plan.RejectedStockQty;
                        item.RejectedSFQty = (int)plan.RejectedSFQty;
                        // Effective available stock accounts for finish stock already allocated/adjusted to this order
                        item.AvailableInHandStock = Math.Max(item.AvailableInHandStock + item.AlreadyAdjustedStockQty, item.StockQty);

                        if (item.OrderQty > planOrderQty)
                        {
                            item.IsSavedPlan = false;
                            item.IsPlanned = true;
                        }
                        else
                        {
                            item.IsSavedPlan = true;
                            item.IsPlanned = true;
                        }
                    }
                    else
                    {
                        item.PreviouslyPlannedOrderQty = 0;
                        item.StockQty = 0;
                        item.RejectedStockQty = 0;
                        item.RejectedSFQty = 0;
                        item.IsSavedPlan = false;
                        item.IsPlanned = false;
                    }

                    // Restore saved SF allocations (only for authorized items)
                    if (item.Authorized)
                    {
                        var itemSavedSf = savedSfAllocations.Where(a => (string)a.ItemID == item.CompItemID).ToList();
                        foreach (var sa in itemSavedSf)
                        {
                            int sooRefId = (int)sa.SOO_RefID;
                            int allocQty = (int)sa.AllocatedQty;
                            int alreadyIssuedQty = (int)sa.AlreadyIssuedQty;
                            var matchingLot = item.AvailableSFLots.FirstOrDefault(l => l.EntryID == sooRefId);
                            int effectiveAvailable = matchingLot != null 
                                ? Math.Max(matchingLot.AvailableQty + allocQty, allocQty)
                                : allocQty;

                            if (matchingLot != null)
                            {
                                matchingLot.AvailableQty = effectiveAvailable;
                            }

                            item.SFStockAllocations.Add(new PpcItemSFAllocationDto
                            {
                                EntryID = (int)sa.EntryID,
                                SOO_RefID = sooRefId,
                                ProcessID = (int)sa.ProcessID,
                                ProcessName = (string)sa.ProcessName,
                                LotNo = (string)sa.LotNo,
                                BatchNo = (string)sa.BatchNo,
                                Shelf_RefID = (int)sa.Shelf_RefID,
                                StoreLocation = (string)sa.StoreLocation,
                                AvailableQty = effectiveAvailable,
                                AllocatedQty = allocQty,
                                AlreadyIssuedQty = alreadyIssuedQty
                            });
                        }
                    }

                    // Restore saved purchases (only for authorized items)
                    var itemPurchases = item.Authorized
                        ? savedPurchases.Where(p => (string)p.ItemID == item.CompItemID).ToList()
                        : new List<dynamic>();
                    foreach (var p in itemPurchases)
                    {
                        var alloc = new PpcItemPurchaseAllocationDto
                        {
                            EntryID = (int)p.EntryID,
                            ProcessID = (int)p.ProcessID,
                            ProcessName = (string)p.ProcessName,
                            VendID = (int?)p.VendID,
                            MakerName = (string?)p.MakerName,
                            PurchaseQty = (int)p.PurchaseQty,
                            PurchaseRate = (double)p.PurchaseRate,
                            ExcessQtyPercentage = p.ExcessQtyPercentage != null ? Convert.ToDouble(p.ExcessQtyPercentage) : 50,
                            VendIssued_RefID = (long?)p.VendIssued_RefID,
                            MasterPONo = (string?)p.MasterPONo,
                            IssuedAt = (DateTime?)p.IssuedAt,
                            IssuedBy = (string?)p.IssuedBy,
                            Status = (string)p.Status,
                            OriginalPlannedQty = (int?)p.OriginalPlannedQty,
                            ShortClosedQty = (int)p.ShortClosedQty,
                            ExceptionNotes = (string?)p.ExceptionNotes,
                            VendIssdDetail_EntryID = (long?)p.VendIssdDetail_EntryID,
                            IssuedQty = (int)p.IssuedQty,
                            ReceivedQty = (int)p.ReceivedQty
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
                                PlannedQty = item.Authorized ? item.ProductionQty : 0,
                                StartDate = DateTime.Today,
                                EndDate = header.DeliveryDate ?? DateTime.Today.AddDays(14)
                            };

                            // Check if previously saved
                            if (item.Authorized)
                            {
                                var savedSched = savedSchedules.FirstOrDefault(s => (string)s.ItemID == item.CompItemID && (string)s.Hub_Name == hubName);
                                if (savedSched != null)
                                {
                                    sched.PlannedQty = item.HasDeltaQty ? item.ProductionQty : (int)savedSched.PlannedQty;
                                    sched.StartDate = (DateTime)savedSched.StartDate;
                                    sched.EndDate = (DateTime)savedSched.EndDate;
                                }
                            }

                            // Attach supervisors for this group and hub
                            sched.Supervisors = groupSupervisors
                                .Where(s => s.GroupID == item.GroupID && s.Hub_Name == hubName)
                                .ToList();

                            item.HubSchedules.Add(sched);
                        }
                    }
                }

                // Header is fully planned only if the order is authorized and all items are authorized & planned (without unallocated delta qty)
                header.IsPlanned = header.Authorized && items.Any() && items.All(i => i.Authorized && i.IsPlanned && !i.HasDeltaQty);

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
                await EnsureExceptionSchemaAsync(db);
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
                await EnsureExceptionSchemaAsync(db);
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
            await EnsureExceptionSchemaAsync(conn);
            using var trans = conn.BeginTransaction();

            try
            {
                var orderNo = request.OrderNo.Trim();
                var user = request.PlannedBy ?? "System";

                // Filter to only items that were actually planned
                var plannedItems = request.Items
                    .Where(i => i.Authorized && (i.IsPlanned || i.HasDeltaQty))
                    .ToList();

                // Server-side floor validations against live downstream execution
                const string liveAdjSql = @"
                    SELECT ItemID, ISNULL(SUM(Qty), 0) AS AlreadyAdjustedQty
                    FROM StockOrderAdjustments
                    WHERE OrderNo = @OrderNo
                    GROUP BY ItemID";
                var liveAdjMap = (await conn.QueryAsync(liveAdjSql, new { OrderNo = orderNo }, trans))
                    .ToDictionary(x => (string)x.ItemID, x => (int)x.AlreadyAdjustedQty);

                const string liveSfSql = @"
                    SELECT VID.ItemCode AS ItemID, SOOI.SOO_RefID, ISNULL(SUM(SOOI.QtyIssd), 0) AS IssuedQty
                    FROM StockOrderOpening_Issuance SOOI
                    INNER JOIN VendIssdDetail VID ON SOOI.VID_RefID = VID.EntryID
                    WHERE VID.OrderNo = @OrderNo
                    GROUP BY VID.ItemCode, SOOI.SOO_RefID";
                var liveSfRows = (await conn.QueryAsync(liveSfSql, new { OrderNo = orderNo }, trans)).ToList();

                foreach (var item in plannedItems)
                {
                    int liveAdjusted = liveAdjMap.TryGetValue(item.CompItemID, out int a) ? a : 0;
                    int minStockFloor = Math.Max(liveAdjusted, item.RejectedStockQty);
                    if (item.StockQty < minStockFloor)
                    {
                        throw new InvalidOperationException($"Finish Stock for {item.CompItemID} cannot be less than already-adjusted/rejected quantity ({minStockFloor} pcs). Use Handle Exception to record rejections.");
                    }

                    foreach (var sf in item.SFStockAllocations)
                    {
                        var liveSfIssued = liveSfRows
                            .Where(r => (string)r.ItemID == item.CompItemID && (int)r.SOO_RefID == sf.SOO_RefID)
                            .Sum(r => (int)r.IssuedQty);
                        if (sf.AllocatedQty < liveSfIssued)
                        {
                            throw new InvalidOperationException($"SF Lot {sf.LotNo} for {item.CompItemID} cannot be reduced below already-issued quantity ({liveSfIssued} pcs).");
                        }
                    }

                    foreach (var p in item.Purchases.Where(x => !string.Equals(x.Status, "Cancelled", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (p.IsIssued && p.PurchaseQty < p.ReceivedQty)
                        {
                            throw new InvalidOperationException($"Maker PO #{p.MasterPONo} for {item.CompItemID} cannot be less than already received quantity ({p.ReceivedQty} pcs).");
                        }
                    }
                }

                int totalAuthorizedCount = request.Items.Count(i => i.Authorized);
                string status = (plannedItems.Count >= totalAuthorizedCount && totalAuthorizedCount > 0)
                    ? "Planned"
                    : (plannedItems.Count > 0 ? "Partially Planned" : "Draft");

                // 1. Delete previous hub schedules, item planning, SF allocations, and master records (Purchases are upserted to preserve issued POs)
                await conn.ExecuteAsync("DELETE FROM PPC_Order_Item_Hub_Schedules WHERE OrderNo = @OrderNo", new { OrderNo = orderNo }, trans);
                await conn.ExecuteAsync("DELETE FROM PPC_Order_Item_SF_Allocations WHERE OrderNo = @OrderNo", new { OrderNo = orderNo }, trans);
                await conn.ExecuteAsync("DELETE FROM PPC_Order_Item_Planning WHERE OrderNo = @OrderNo", new { OrderNo = orderNo }, trans);
                await conn.ExecuteAsync("DELETE FROM PPC_Order_Planning_Master WHERE OrderNo = @OrderNo", new { OrderNo = orderNo }, trans);

                // Fetch existing purchase allocations to protect already-issued Master POs
                const string existingPurchasesSql = @"
                    SELECT EntryID, ItemID, ProcessID, VendID, PurchaseQty, PurchaseRate, ExcessQtyPercentage, VendIssued_RefID, MasterPONo
                    FROM PPC_Order_Item_Purchases
                    WHERE OrderNo = @OrderNo";
                var existingPurchases = (await conn.QueryAsync<dynamic>(existingPurchasesSql, new { OrderNo = orderNo }, trans)).ToList();
                var retainedEntryIds = new List<int>();

                // 2. Insert Master record
                const string masterSql = @"
                    INSERT INTO PPC_Order_Planning_Master (OrderNo, PlannedBy, PlannedAt, Status, Notes)
                    VALUES (@OrderNo, @PlannedBy, GETUTCDATE(), @Status, @Notes)";

                await conn.ExecuteAsync(masterSql, new
                {
                    OrderNo = orderNo,
                    PlannedBy = user,
                    Status = status,
                    Notes = request.Notes
                }, trans);

                // 3. Insert Item Planning & Upsert Purchases & Hub Schedules (only for planned items)
                const string itemPlanSql = @"
                    INSERT INTO PPC_Order_Item_Planning 
                        (OrderNo, ItemID, OrderQty, StockQty, SFStockQty, RejectedStockQty, RejectedSFQty, TotalPurchaseQty, ProductionQty, Remarks, CreatedAt, CreatedBy)
                    VALUES 
                        (@OrderNo, @ItemID, @OrderQty, @StockQty, @SFStockQty, @RejectedStockQty, @RejectedSFQty, @TotalPurchaseQty, @ProductionQty, @Remarks, GETUTCDATE(), @CreatedBy)";

                const string insertSfAllocSql = @"
                    INSERT INTO PPC_Order_Item_SF_Allocations
                        (OrderNo, ItemID, SOO_RefID, ProcessID, ProcessName, LotNo, BatchNo, Shelf_RefID, StoreLocation, AllocatedQty, CreatedAt, CreatedBy)
                    VALUES
                        (@OrderNo, @ItemID, @SOO_RefID, @ProcessID, @ProcessName, @LotNo, @BatchNo, @Shelf_RefID, @StoreLocation, @AllocatedQty, GETUTCDATE(), @CreatedBy)";

                const string insertPurchaseSql = @"
                    INSERT INTO PPC_Order_Item_Purchases 
                        (OrderNo, ItemID, ProcessID, VendID, PurchaseQty, PurchaseRate, ExcessQtyPercentage, CreatedAt, CreatedBy)
                    VALUES 
                        (@OrderNo, @ItemID, @ProcessID, @VendID, @PurchaseQty, @PurchaseRate, @ExcessQtyPercentage, GETUTCDATE(), @CreatedBy);
                    SELECT SCOPE_IDENTITY();";

                const string updatePurchaseSql = @"
                    UPDATE PPC_Order_Item_Purchases 
                    SET ProcessID = @ProcessID,
                        VendID = @VendID,
                        PurchaseQty = @PurchaseQty,
                        PurchaseRate = @PurchaseRate,
                        ExcessQtyPercentage = @ExcessQtyPercentage
                    WHERE EntryID = @EntryID AND (VendIssued_RefID IS NULL OR VendIssued_RefID = 0)";

                const string schedSql = @"
                    INSERT INTO PPC_Order_Item_Hub_Schedules 
                        (OrderNo, ItemID, Hub_Name, PlannedQty, StartDate, EndDate, CreatedAt, CreatedBy)
                    VALUES 
                        (@OrderNo, @ItemID, @Hub_Name, @PlannedQty, @StartDate, @EndDate, GETUTCDATE(), @CreatedBy)";

                foreach (var item in plannedItems)
                {
                    // Item summary row
                    await conn.ExecuteAsync(itemPlanSql, new
                    {
                        OrderNo = orderNo,
                        ItemID = item.CompItemID,
                        OrderQty = item.OrderQty,
                        StockQty = item.StockQty,
                        SFStockQty = item.SFStockQty,
                        RejectedStockQty = item.RejectedStockQty,
                        RejectedSFQty = item.RejectedSFQty,
                        TotalPurchaseQty = item.TotalPurchaseQty,
                        ProductionQty = item.ProductionQty,
                        Remarks = "",
                        CreatedBy = user
                    }, trans);

                    // Semi-finished stock allocations
                    foreach (var sfAlloc in item.SFStockAllocations.Where(a => a.AllocatedQty > 0))
                    {
                        await conn.ExecuteAsync(insertSfAllocSql, new
                        {
                            OrderNo = orderNo,
                            ItemID = item.CompItemID,
                            sfAlloc.SOO_RefID,
                            sfAlloc.ProcessID,
                            sfAlloc.ProcessName,
                            sfAlloc.LotNo,
                            sfAlloc.BatchNo,
                            sfAlloc.Shelf_RefID,
                            sfAlloc.StoreLocation,
                            sfAlloc.AllocatedQty,
                            CreatedBy = user
                        }, trans);
                    }

                    // Multi-purchase split rows: upsert without breaking already-issued POs
                    foreach (var p in item.Purchases.Where(x => x.PurchaseQty > 0 && x.ProcessID > 0))
                    {
                        var existing = existingPurchases.FirstOrDefault(e => (int)e.EntryID == p.EntryID);
                        if (existing != null)
                        {
                            retainedEntryIds.Add(p.EntryID);
                            bool isIssued = existing.VendIssued_RefID != null && (long)existing.VendIssued_RefID > 0;
                            if (!isIssued)
                            {
                                await conn.ExecuteAsync(updatePurchaseSql, new
                                {
                                    ProcessID = p.ProcessID,
                                    VendID = p.VendID,
                                    PurchaseQty = p.PurchaseQty,
                                    PurchaseRate = p.PurchaseRate,
                                    ExcessQtyPercentage = p.ExcessQtyPercentage > 0 ? p.ExcessQtyPercentage : 50,
                                    EntryID = p.EntryID
                                }, trans);
                            }
                            // If already issued, leave strictly untouched to preserve PO linkage and integrity
                        }
                        else
                        {
                            int newEntryId = await conn.ExecuteScalarAsync<int>(insertPurchaseSql, new
                            {
                                OrderNo = orderNo,
                                ItemID = item.CompItemID,
                                ProcessID = p.ProcessID,
                                VendID = p.VendID,
                                PurchaseQty = p.PurchaseQty,
                                PurchaseRate = p.PurchaseRate,
                                ExcessQtyPercentage = p.ExcessQtyPercentage > 0 ? p.ExcessQtyPercentage : 50,
                                CreatedBy = user
                            }, trans);

                            p.EntryID = newEntryId;
                            retainedEntryIds.Add(newEntryId);
                        }
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

                // Delete only UNISSUED active purchase allocations that were removed from the plan
                if (retainedEntryIds.Any())
                {
                    const string deleteUnretainedSql = @"
                        DELETE FROM PPC_Order_Item_Purchases 
                        WHERE OrderNo = @OrderNo 
                          AND ISNULL(Status, 'Active') = 'Active'
                          AND (VendIssued_RefID IS NULL OR VendIssued_RefID = 0)
                          AND EntryID NOT IN @RetainedEntryIds";
                    await conn.ExecuteAsync(deleteUnretainedSql, new { OrderNo = orderNo, RetainedEntryIds = retainedEntryIds }, trans);
                }
                else
                {
                    await conn.ExecuteAsync("DELETE FROM PPC_Order_Item_Purchases WHERE OrderNo = @OrderNo AND ISNULL(Status, 'Active') = 'Active' AND (VendIssued_RefID IS NULL OR VendIssued_RefID = 0)", new { OrderNo = orderNo }, trans);
                }

                trans.Commit();
                _logger.LogInformation("PPC Plan successfully saved for Order #{OrderNo} with {ItemCount} planned items (Status: {Status}).", orderNo, plannedItems.Count, status);
                return true;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                _logger.LogError(ex, "Failed to save PPC Plan for Order #{OrderNo}", request.OrderNo);
                throw;
            }
        }

        public async Task<PpcExceptionActionResult> ExecutePlanningExceptionAsync(PpcExceptionActionRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.OrderNo) || string.IsNullOrWhiteSpace(request.ItemID))
                return new PpcExceptionActionResult { Success = false, Message = "Order and Item are required." };

            if (string.IsNullOrWhiteSpace(request.Reason))
                return new PpcExceptionActionResult { Success = false, Message = "A mandatory reason is required for audit tracking." };

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            await EnsureExceptionSchemaAsync(conn);
            using var trans = conn.BeginTransaction();

            try
            {
                string orderNo = request.OrderNo.Trim();
                string itemId = request.ItemID.Trim();
                string actionBy = string.IsNullOrWhiteSpace(request.ActionBy) ? "PPC Planner" : request.ActionBy.Trim();
                string reason = request.Reason.Trim();
                string? sourceRef = null;
                int affectedQty = request.AffectedQty;
                string resultMessage = "Exception recorded and item opened for re-planning.";

                switch (request.ExceptionType)
                {
                    case "FinishStockRelease":
                    {
                        var plan = await conn.QueryFirstOrDefaultAsync<dynamic>(
                            "SELECT StockQty FROM PPC_Order_Item_Planning WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { OrderNo = orderNo, ItemID = itemId }, trans);
                        if (plan == null) throw new InvalidOperationException("Item planning record not found.");

                        int stockQty = (int)plan.StockQty;
                        int alreadyAdjusted = await conn.ExecuteScalarAsync<int>(
                            "SELECT ISNULL(SUM(Qty), 0) FROM StockOrderAdjustments WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { OrderNo = orderNo, ItemID = itemId }, trans);

                        int unissued = Math.Max(0, stockQty - alreadyAdjusted);
                        if (affectedQty <= 0 || affectedQty > unissued)
                            throw new InvalidOperationException($"Release quantity ({affectedQty}) must be between 1 and {unissued} unissued pcs.");

                        await conn.ExecuteAsync(
                            @"UPDATE PPC_Order_Item_Planning 
                              SET StockQty = StockQty - @Qty, ProductionQty = ProductionQty + @Qty 
                              WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { Qty = affectedQty, OrderNo = orderNo, ItemID = itemId }, trans);

                        sourceRef = "Finish Stock";
                        resultMessage = $"Released {affectedQty} pcs of unissued Finish Stock for re-planning.";
                        break;
                    }

                    case "FinishStockRejection":
                    {
                        var plan = await conn.QueryFirstOrDefaultAsync<dynamic>(
                            "SELECT StockQty, ISNULL(RejectedStockQty, 0) AS RejectedStockQty FROM PPC_Order_Item_Planning WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { OrderNo = orderNo, ItemID = itemId }, trans);
                        if (plan == null) throw new InvalidOperationException("Item planning record not found.");

                        int rejectedSoFar = (int)plan.RejectedStockQty;
                        int alreadyAdjusted = await conn.ExecuteScalarAsync<int>(
                            "SELECT ISNULL(SUM(Qty), 0) FROM StockOrderAdjustments WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { OrderNo = orderNo, ItemID = itemId }, trans);

                        int maxRejectable = Math.Max(0, alreadyAdjusted - rejectedSoFar);
                        if (affectedQty <= 0 || affectedQty > maxRejectable)
                            throw new InvalidOperationException($"Rejection quantity ({affectedQty}) must be between 1 and {maxRejectable} pcs already withdrawn from store.");

                        await conn.ExecuteAsync(
                            @"UPDATE PPC_Order_Item_Planning 
                              SET RejectedStockQty = ISNULL(RejectedStockQty, 0) + @Qty, ProductionQty = ProductionQty + @Qty 
                              WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { Qty = affectedQty, OrderNo = orderNo, ItemID = itemId }, trans);

                        sourceRef = "Finish Stock QC Rejection";
                        resultMessage = $"Recorded {affectedQty} pcs Finish Stock QC rejection. Replacement quantity opened for re-planning.";
                        break;
                    }

                    case "SFStockRelease":
                    {
                        if (!request.SfAllocationSooRefID.HasValue || request.SfAllocationSooRefID.Value <= 0)
                            throw new InvalidOperationException("Please select the SF Lot allocation.");

                        int sooRefId = request.SfAllocationSooRefID.Value;
                        var sfAlloc = await conn.QueryFirstOrDefaultAsync<dynamic>(
                            "SELECT EntryID, LotNo, AllocatedQty FROM PPC_Order_Item_SF_Allocations WHERE OrderNo = @OrderNo AND ItemID = @ItemID AND SOO_RefID = @SooRefID",
                            new { OrderNo = orderNo, ItemID = itemId, SooRefID = sooRefId }, trans);
                        if (sfAlloc == null) throw new InvalidOperationException("SF lot allocation record not found.");

                        int allocated = (int)sfAlloc.AllocatedQty;
                        int alreadyIssued = await conn.ExecuteScalarAsync<int>(
                            @"SELECT ISNULL(SUM(SOOI.QtyIssd), 0)
                              FROM StockOrderOpening_Issuance SOOI
                              INNER JOIN VendIssdDetail VID ON SOOI.VID_RefID = VID.EntryID
                              WHERE VID.OrderNo = @OrderNo AND SOOI.SOO_RefID = @SooRefID",
                            new { OrderNo = orderNo, SooRefID = sooRefId }, trans);

                        int unissued = Math.Max(0, allocated - alreadyIssued);
                        if (affectedQty <= 0 || affectedQty > unissued)
                            throw new InvalidOperationException($"SF release quantity ({affectedQty}) must be between 1 and {unissued} unissued pcs.");

                        int remainingAlloc = allocated - affectedQty;
                        if (remainingAlloc <= 0)
                        {
                            await conn.ExecuteAsync("DELETE FROM PPC_Order_Item_SF_Allocations WHERE EntryID = @EntryID", new { EntryID = (int)sfAlloc.EntryID }, trans);
                        }
                        else
                        {
                            await conn.ExecuteAsync("UPDATE PPC_Order_Item_SF_Allocations SET AllocatedQty = @Remaining WHERE EntryID = @EntryID",
                                new { Remaining = remainingAlloc, EntryID = (int)sfAlloc.EntryID }, trans);
                        }

                        await conn.ExecuteAsync(
                            @"UPDATE PPC_Order_Item_Planning 
                              SET SFStockQty = CASE WHEN ISNULL(SFStockQty, 0) >= @Qty THEN ISNULL(SFStockQty, 0) - @Qty ELSE 0 END,
                                  ProductionQty = ProductionQty + @Qty
                              WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { Qty = affectedQty, OrderNo = orderNo, ItemID = itemId }, trans);

                        sourceRef = $"SF Lot {sfAlloc.LotNo}";
                        resultMessage = $"Released {affectedQty} pcs from SF Lot {sfAlloc.LotNo} for re-planning.";
                        break;
                    }

                    case "SFStockRejection":
                    {
                        var plan = await conn.QueryFirstOrDefaultAsync<dynamic>(
                            "SELECT ISNULL(SFStockQty, 0) AS SFStockQty, ISNULL(RejectedSFQty, 0) AS RejectedSFQty FROM PPC_Order_Item_Planning WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { OrderNo = orderNo, ItemID = itemId }, trans);
                        if (plan == null) throw new InvalidOperationException("Item planning record not found.");

                        int totalSfIssued = await conn.ExecuteScalarAsync<int>(
                            @"SELECT ISNULL(SUM(SOOI.QtyIssd), 0)
                              FROM StockOrderOpening_Issuance SOOI
                              INNER JOIN VendIssdDetail VID ON SOOI.VID_RefID = VID.EntryID
                              WHERE VID.OrderNo = @OrderNo AND VID.ItemCode = @ItemID",
                            new { OrderNo = orderNo, ItemID = itemId }, trans);

                        int rejectedSf = (int)plan.RejectedSFQty;
                        int sfCeiling = Math.Max(totalSfIssued, (int)plan.SFStockQty) - rejectedSf;
                        if (affectedQty <= 0 || affectedQty > sfCeiling)
                            throw new InvalidOperationException($"SF rejection quantity ({affectedQty}) must be between 1 and {sfCeiling} pcs.");

                        await conn.ExecuteAsync(
                            @"UPDATE PPC_Order_Item_Planning 
                              SET RejectedSFQty = ISNULL(RejectedSFQty, 0) + @Qty, ProductionQty = ProductionQty + @Qty 
                              WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { Qty = affectedQty, OrderNo = orderNo, ItemID = itemId }, trans);

                        sourceRef = "SF Stock QC Rejection";
                        resultMessage = $"Recorded {affectedQty} pcs SF Stock rejection. Replacement quantity opened for re-planning.";
                        break;
                    }

                    case "CancelMakerPO":
                    {
                        if (!request.PurchaseSplitEntryID.HasValue || request.PurchaseSplitEntryID.Value <= 0)
                            throw new InvalidOperationException("Please select the Maker PO allocation to cancel.");

                        var purch = await conn.QueryFirstOrDefaultAsync<dynamic>(
                            @"SELECT EntryID, ProcessID, PurchaseQty, VendIssued_RefID, MasterPONo, ISNULL(Status, 'Active') AS Status
                              FROM PPC_Order_Item_Purchases
                              WHERE EntryID = @EntryID AND OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { EntryID = request.PurchaseSplitEntryID.Value, OrderNo = orderNo, ItemID = itemId }, trans);
                        if (purch == null) throw new InvalidOperationException("Purchase split not found.");
                        if (string.Equals((string)purch.Status, "Cancelled", StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("This purchase split is already cancelled.");

                        int purchaseQty = (int)purch.PurchaseQty;
                        long? vendIssuedRefId = (long?)purch.VendIssued_RefID;
                        string masterPoNo = (string?)purch.MasterPONo ?? string.Empty;

                        if (vendIssuedRefId.HasValue && vendIssuedRefId.Value > 0)
                        {
                            var detailRow = await conn.QueryFirstOrDefaultAsync<dynamic>(
                                @"SELECT TOP 1 EntryID, CAST(ISNULL(IssQty, 0) AS INT) AS IssQty
                                  FROM VendIssdDetail
                                  WHERE RefID = @RefID AND OrderNo = @OrderNo AND ItemCode = @ItemID AND RcvProcessID = @ProcessID
                                  ORDER BY EntryID DESC",
                                new { RefID = vendIssuedRefId.Value, OrderNo = orderNo, ItemID = itemId, ProcessID = (int)purch.ProcessID }, trans);

                            if (detailRow != null)
                            {
                                long detailEntryId = (long)detailRow.EntryID;
                                int rcvdQty = await conn.ExecuteScalarAsync<int>(
                                    "SELECT CAST(ISNULL(SUM(RcvdQty), 0) AS INT) FROM VendRcvdDetail WHERE Issue_RefID = @DetailEntryID",
                                    new { DetailEntryID = detailEntryId }, trans);

                                if (rcvdQty > 0)
                                    throw new InvalidOperationException($"PO #{masterPoNo} already has {rcvdQty} pcs received. Use 'Short-Close PO' instead of Cancel.");

                                await conn.ExecuteAsync("DELETE FROM VendIssdDetail_ReturnDTs WHERE VIS_RefID = @DetailEntryID", new { DetailEntryID = detailEntryId }, trans);
                                await conn.ExecuteAsync("DELETE FROM VendIssdDetail_MoreDetails WHERE VID_RefID = @DetailEntryID", new { DetailEntryID = detailEntryId }, trans);
                                await conn.ExecuteAsync("DELETE FROM VendIssAuthDetail WHERE IssID = @DetailEntryID", new { DetailEntryID = detailEntryId }, trans);
                                await conn.ExecuteAsync("DELETE FROM VendIssdDetail WHERE EntryID = @DetailEntryID", new { DetailEntryID = detailEntryId }, trans);

                                int remainingLines = await conn.ExecuteScalarAsync<int>(
                                    "SELECT COUNT(1) FROM VendIssdDetail WHERE RefID = @RefID",
                                    new { RefID = vendIssuedRefId.Value }, trans);
                                if (remainingLines == 0)
                                {
                                    await conn.ExecuteAsync("DELETE FROM VendIssued WHERE EntryID = @RefID", new { RefID = vendIssuedRefId.Value }, trans);
                                }
                            }
                        }

                        await conn.ExecuteAsync(
                            @"UPDATE PPC_Order_Item_Purchases
                              SET Status = 'Cancelled',
                                  OriginalPlannedQty = ISNULL(OriginalPlannedQty, PurchaseQty),
                                  ExceptionNotes = @Reason,
                                  UpdatedAt = GETUTCDATE(),
                                  UpdatedBy = @ActionBy
                              WHERE EntryID = @EntryID",
                            new { Reason = reason, ActionBy = actionBy, EntryID = (int)purch.EntryID }, trans);

                        await conn.ExecuteAsync(
                            @"UPDATE PPC_Order_Item_Planning
                              SET TotalPurchaseQty = CASE WHEN ISNULL(TotalPurchaseQty, 0) >= @Qty THEN TotalPurchaseQty - @Qty ELSE 0 END,
                                  ProductionQty = ProductionQty + @Qty
                              WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { Qty = purchaseQty, OrderNo = orderNo, ItemID = itemId }, trans);

                        affectedQty = purchaseQty;
                        sourceRef = string.IsNullOrWhiteSpace(masterPoNo) ? $"Split #{purch.EntryID}" : $"PO #{masterPoNo}";
                        resultMessage = $"Cancelled {sourceRef} ({purchaseQty} pcs). Full quantity freed for re-planning.";
                        break;
                    }

                    case "ShortCloseMakerPO":
                    {
                        if (!request.PurchaseSplitEntryID.HasValue || request.PurchaseSplitEntryID.Value <= 0)
                            throw new InvalidOperationException("Please select the Maker PO allocation to short-close.");

                        var purch = await conn.QueryFirstOrDefaultAsync<dynamic>(
                            @"SELECT EntryID, ProcessID, PurchaseQty, VendIssued_RefID, MasterPONo, ISNULL(Status, 'Active') AS Status
                              FROM PPC_Order_Item_Purchases
                              WHERE EntryID = @EntryID AND OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { EntryID = request.PurchaseSplitEntryID.Value, OrderNo = orderNo, ItemID = itemId }, trans);
                        if (purch == null) throw new InvalidOperationException("Purchase split not found.");
                        if (string.Equals((string)purch.Status, "Cancelled", StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("Cannot short-close a cancelled PO split.");

                        int purchaseQty = (int)purch.PurchaseQty;
                        long? vendIssuedRefId = (long?)purch.VendIssued_RefID;
                        string masterPoNo = (string?)purch.MasterPONo ?? string.Empty;

                        if (!vendIssuedRefId.HasValue || vendIssuedRefId.Value <= 0)
                            throw new InvalidOperationException("This split has not been issued as a PO yet. Unlock and edit/delete it instead.");

                        var detailRow = await conn.QueryFirstOrDefaultAsync<dynamic>(
                            @"SELECT TOP 1 EntryID
                              FROM VendIssdDetail
                              WHERE RefID = @RefID AND OrderNo = @OrderNo AND ItemCode = @ItemID AND RcvProcessID = @ProcessID
                              ORDER BY EntryID DESC",
                            new { RefID = vendIssuedRefId.Value, OrderNo = orderNo, ItemID = itemId, ProcessID = (int)purch.ProcessID }, trans);
                        if (detailRow == null) throw new InvalidOperationException("Linked VendIssdDetail PO line not found.");

                        long detailEntryId = (long)detailRow.EntryID;
                        int rcvdQty = await conn.ExecuteScalarAsync<int>(
                            "SELECT CAST(ISNULL(SUM(RcvdQty), 0) AS INT) FROM VendRcvdDetail WHERE Issue_RefID = @DetailEntryID",
                            new { DetailEntryID = detailEntryId }, trans);

                        if (rcvdQty <= 0)
                            throw new InvalidOperationException($"PO #{masterPoNo} has 0 received pcs. Use 'Cancel Maker PO' instead.");
                        if (rcvdQty >= purchaseQty)
                            throw new InvalidOperationException($"PO #{masterPoNo} is already fully received ({rcvdQty}/{purchaseQty} pcs).");

                        int shortfallQty = purchaseQty - rcvdQty;

                        await conn.ExecuteAsync(
                            "UPDATE VendIssdDetail SET IssQty = @RcvdQty WHERE EntryID = @DetailEntryID",
                            new { RcvdQty = rcvdQty, DetailEntryID = detailEntryId }, trans);
                        await conn.ExecuteAsync(
                            "UPDATE VendIssdDetail_ReturnDTs SET IssQty = @RcvdQty WHERE VIS_RefID = @DetailEntryID",
                            new { RcvdQty = rcvdQty, DetailEntryID = detailEntryId }, trans);

                        await conn.ExecuteAsync(
                            @"UPDATE PPC_Order_Item_Purchases
                              SET OriginalPlannedQty = ISNULL(OriginalPlannedQty, PurchaseQty),
                                  PurchaseQty = @RcvdQty,
                                  ShortClosedQty = ISNULL(ShortClosedQty, 0) + @ShortfallQty,
                                  Status = 'ShortClosed',
                                  ExceptionNotes = @Reason,
                                  UpdatedAt = GETUTCDATE(),
                                  UpdatedBy = @ActionBy
                              WHERE EntryID = @EntryID",
                            new { RcvdQty = rcvdQty, ShortfallQty = shortfallQty, Reason = reason, ActionBy = actionBy, EntryID = (int)purch.EntryID }, trans);

                        await conn.ExecuteAsync(
                            @"UPDATE PPC_Order_Item_Planning
                              SET TotalPurchaseQty = CASE WHEN ISNULL(TotalPurchaseQty, 0) >= @ShortfallQty THEN TotalPurchaseQty - @ShortfallQty ELSE 0 END,
                                  ProductionQty = ProductionQty + @ShortfallQty
                              WHERE OrderNo = @OrderNo AND ItemID = @ItemID",
                            new { ShortfallQty = shortfallQty, OrderNo = orderNo, ItemID = itemId }, trans);

                        affectedQty = shortfallQty;
                        sourceRef = string.IsNullOrWhiteSpace(masterPoNo) ? $"Split #{purch.EntryID}" : $"PO #{masterPoNo}";
                        resultMessage = $"Short-closed {sourceRef} at received qty ({rcvdQty} pcs). Remaining {shortfallQty} pcs freed for re-planning.";
                        break;
                    }

                    case "UnlockPlan":
                    {
                        sourceRef = "Manual Plan Revision";
                        resultMessage = "Item unlocked for allocation revision.";
                        break;
                    }

                    default:
                        throw new InvalidOperationException($"Unsupported exception type: {request.ExceptionType}");
                }

                await conn.ExecuteAsync(
                    @"INSERT INTO PPC_Order_Planning_Exceptions (OrderNo, ItemID, ExceptionType, AffectedQty, SourceRef, Reason, CreatedBy, CreatedAt)
                      VALUES (@OrderNo, @ItemID, @ExceptionType, @AffectedQty, @SourceRef, @Reason, @CreatedBy, GETUTCDATE())",
                    new
                    {
                        OrderNo = orderNo,
                        ItemID = itemId,
                        request.ExceptionType,
                        AffectedQty = affectedQty,
                        SourceRef = sourceRef,
                        Reason = reason,
                        CreatedBy = actionBy
                    }, trans);

                trans.Commit();
                return new PpcExceptionActionResult { Success = true, Message = resultMessage };
            }
            catch (Exception ex)
            {
                trans.Rollback();
                _logger.LogError(ex, "Error executing PPC planning exception {ExceptionType} for Order #{OrderNo} Item {ItemID}",
                    request.ExceptionType, request.OrderNo, request.ItemID);
                return new PpcExceptionActionResult { Success = false, Message = ex.Message };
            }
        }
    }
}
