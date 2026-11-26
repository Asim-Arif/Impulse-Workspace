-- ============================================================================
-- Production Item List Performance: Indexes, Optimized VRunningLots_POI, & SP_GetProductionItemList
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_POI_Remarks_LotNo' AND object_id = OBJECT_ID(N'dbo.POI_Remarks'))
    CREATE NONCLUSTERED INDEX IX_POI_Remarks_LotNo
    ON dbo.POI_Remarks (LotNo) INCLUDE (Remarks);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_POI_Remarks_OrderNo_ItemCode' AND object_id = OBJECT_ID(N'dbo.POI_Remarks'))
    CREATE NONCLUSTERED INDEX IX_POI_Remarks_OrderNo_ItemCode
    ON dbo.POI_Remarks (OrderNo, ItemCode) INCLUDE (Remarks);
GO

CREATE OR ALTER VIEW [dbo].[VRunningLots_POI]
AS
SELECT
    TUnion.StartDT,
    TUnion.LotNo,
    TUnion.ItemCode,
    TUnion.Qty,
    TUnion.Description,
    TUnion.LastProcessDays,
    TUnion.OrderNo,
    TUnion.Remarks,
    TUnion.DT,
    TUnion.ProcessID,
    TUnion.ReWorkLot,
    Items_1.GroupID,
    FCustomerOrders_1.CustCode,
    FCustomerOrders_1.Country,
    TRunning.TLRunning,
    TRunning.TLRunning AS TLRunning_Order,
    TSF.R1,
    TSF.R2,
    TPack.PackStock,
    Items_1.ItemName,
    VUnshippedOrderItemList_1.OrderNo AS UnShipped_OrderNo,
    TUnion.EntryType,
    TUnion.POIRemarks,
    TUnion.OpenPOQty,
    TUnion.VendID,
    TUnion.R1 AS RunningLots_R1,
    TLOrder.TLUnShippedQty,
    TLot_Wastage.Lot_Wastage_Qty,
    VUnshippedOrderItemList_1.Price,
    VUnshippedOrderItemList_1.PriceForCost,
    FOrderItems.Price AS FOI_Price,
    dbo.ForeignCustomers.Curr,
    TTray.TrayNo,
    TUnion.MasterPONo,
    FOrderItems.DeliveryDT AS ItemDeliveryDT,
    dbo.ProcessGroupsProcesses.Hub_Name,
    TUnion.TrayNo AS Lot_TrayNo,
    TDispatch.Qty AS DispatchQty,
    FOrderItems.OrderQty
FROM
(
    SELECT
        ISNULL(T1.StartDT, T2.StartDT) AS StartDT,
        VRL.LotNo,
        VRL.ItemCode,
        VRL.Qty,
        VRL.Description,
        DATEDIFF(day, VRL.DT, GETDATE()) AS LastProcessDays,
        VRL.OrderNo,
        '' AS Remarks,
        VRL.DT,
        VRL.ProcessID,
        VRL.ReWorkLot,
        0 AS EntryType,
        dbo.POI_Remarks.Remarks AS POIRemarks,
        0 AS OpenPOQty,
        VRL.VendID,
        CASE WHEN VRL.SNO <= VRL.PolishingItemSNo THEN 1 ELSE 0 END AS R1,
        dbo.VLotsWithMasterPONo.MasterPONo,
        QC.TrayNo
    FROM dbo.VRunningLots_Simple AS VRL WITH (NOLOCK)
    INNER JOIN dbo.Items AS I_RL ON VRL.ItemCode = I_RL.ItemID
    LEFT OUTER JOIN dbo.VendRcvdDetail_QC_Report AS QC WITH (NOLOCK) ON VRL.VRD_EntryID = QC.VRD_RefID
    LEFT OUTER JOIN dbo.POI_Remarks ON VRL.LotNo = dbo.POI_Remarks.LotNo
    LEFT OUTER JOIN dbo.VLotsWithMasterPONo ON VRL.LotNo = dbo.VLotsWithMasterPONo.LotNo
    OUTER APPLY (
        SELECT MIN(vr.DT) AS StartDT
        FROM dbo.vendRcvdDetail AS vrd
        INNER JOIN dbo.VendReceived AS vr ON vr.EntryID = vrd.RefID
        WHERE vrd.LotNo = VRL.LotNo
    ) AS T1
    OUTER APPLY (
        SELECT MIN(vi.DT) AS StartDT
        FROM dbo.VendIssdDetail AS vid
        INNER JOIN dbo.VendIssued AS vi ON vi.EntryID = vid.RefID
        WHERE vid.LotNo = VRL.LotNo
    ) AS T2

    UNION ALL

    SELECT
        u.DT,
        '' AS LotNo,
        u.CompItemCode,
        u.OrderQty - u.ShippedQty - ISNULL(T1_1.RLQty, 0) AS Qty,
        '' AS Description,
        0 AS LastProcessDays,
        u.OrderNo,
        u.Remarks,
        u.DT AS Expr1,
        0 AS ProcessID,
        0 AS ReWorkLot,
        1 AS EntryType,
        POI_Remarks_1.Remarks AS POIRemarks,
        TOpenPOs.OpenPOQty,
        0 AS VendID,
        0 AS R1,
        TOpenPOs.MasterPONo,
        '' AS TrayNo
    FROM dbo.VUnshippedOrderItemList AS u
    LEFT OUTER JOIN dbo.POI_Remarks AS POI_Remarks_1
        ON u.OrderNo = POI_Remarks_1.OrderNo
       AND u.CompItemCode = POI_Remarks_1.ItemCode
    OUTER APPLY (
        SELECT SUM(op.Qty) AS OpenPOQty, MAX(op.MasterPONo) AS MasterPONo
        FROM dbo.VItemOpenPOs AS op
        WHERE op.OrderNo = u.OrderNo
          AND op.ItemCode = u.CompItemCode
    ) AS TOpenPOs
    OUTER APPLY (
        SELECT SUM(vrl3.Qty) AS RLQty
        FROM dbo.VRunningLots_Simple AS vrl3 WITH (NOLOCK)
        WHERE vrl3.OrderNo = u.OrderNo
          AND vrl3.ItemCode = u.CompItemCode
    ) AS T1_1
) AS TUnion
INNER JOIN dbo.Items AS Items_1 ON TUnion.ItemCode = Items_1.ItemID
INNER JOIN (
    SELECT OrderNo, CompItemCode, MAX(Price) AS Price, MAX(DeliveryDT) AS DeliveryDT, SUM(Qty) AS OrderQty
    FROM dbo.FOrderItems
    GROUP BY OrderNo, CompItemCode
) AS FOrderItems ON TUnion.ItemCode = FOrderItems.CompItemCode AND TUnion.OrderNo = FOrderItems.OrderNo
LEFT OUTER JOIN dbo.FCustomerOrders AS FCustomerOrders_1 ON FCustomerOrders_1.OrderNo = TUnion.OrderNo
LEFT OUTER JOIN dbo.ForeignCustomers
    ON dbo.ForeignCustomers.CustCode = FCustomerOrders_1.CustCode
   AND dbo.ForeignCustomers.Country = FCustomerOrders_1.Country
LEFT OUTER JOIN dbo.VUnshippedOrderItemList AS VUnshippedOrderItemList_1
    ON TUnion.ItemCode = VUnshippedOrderItemList_1.CompItemCode
   AND TUnion.OrderNo = VUnshippedOrderItemList_1.OrderNo
LEFT OUTER JOIN dbo.ItemProcessGroups ON TUnion.ItemCode = dbo.ItemProcessGroups.ItemID
LEFT OUTER JOIN dbo.ProcessGroupsProcesses
    ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroupsProcesses.Group_RefID
   AND TUnion.ProcessID = dbo.ProcessGroupsProcesses.Process_RefID
OUTER APPLY (
    SELECT SUM(vrl2.Qty) AS TLRunning
    FROM dbo.VRunningLots_Simple AS vrl2 WITH (NOLOCK)
    WHERE vrl2.OrderNo = TUnion.OrderNo
      AND vrl2.ItemCode = TUnion.ItemCode
) AS TRunning
OUTER APPLY (
    SELECT
        SUM(CASE WHEN vso.R1 = 1 THEN vso.Qty - vso.IssdQty ELSE 0 END) AS R1,
        SUM(CASE WHEN vso.R1 = 0 THEN vso.Qty - vso.IssdQty ELSE 0 END) AS R2
    FROM dbo.VStockOrderOpening AS vso
    WHERE vso.Qty > vso.IssdQty
      AND vso.ItemID = TUnion.ItemCode
) AS TSF
OUTER APPLY (
    SELECT SUM(ws.NetQty) AS PackStock
    FROM dbo.VItemsWithShelfWiseStock AS ws
    WHERE ws.ItemID = TUnion.ItemCode
) AS TPack
OUTER APPLY (
    SELECT SUM(u2.OrderQty - u2.ShippedQty) AS TLUnShippedQty
    FROM dbo.VUnshippedOrderItemList AS u2
    WHERE u2.CompItemCode = Items_1.ItemID
) AS TLOrder
OUTER APPLY (
    SELECT SUM(vrd1.Wastage) AS Lot_Wastage_Qty
    FROM dbo.vendRcvdDetail AS vrd1
    WHERE TUnion.LotNo <> ''
      AND vrd1.LotNo = TUnion.LotNo
) AS TLot_Wastage
OUTER APPLY (
    SELECT TOP (1) qc2.TrayNo
    FROM dbo.VendRcvdDetail_QC_Report AS qc2 WITH (NOLOCK)
    INNER JOIN dbo.vendRcvdDetail AS vrd2 ON vrd2.EntryID = qc2.VRD_RefID
    WHERE TUnion.LotNo <> ''
      AND vrd2.LotNo = TUnion.LotNo
) AS TTray
OUTER APPLY (
    SELECT SUM(dl.InnerQty) AS Qty
    FROM dbo.VDispatchListDetail_Ex_Orderwise AS dl
    WHERE dl.OrderNo = TUnion.OrderNo
      AND dl.ItemCode = TUnion.ItemCode
) AS TDispatch
WHERE (VUnshippedOrderItemList_1.OrderNo IS NOT NULL OR TUnion.OrderNo IN ('Stock-Order', 'Stock-OrderIss'))
GO

CREATE OR ALTER PROCEDURE [dbo].[SP_GetProductionItemList]
(
    @LotNo             VARCHAR(50)  = NULL,
    @DateRangeMode     INT          = 0,
    @DtFrom            DATETIME     = NULL,
    @DtTo              DATETIME     = NULL,
    @CustCode          VARCHAR(50)  = NULL,
    @Country           VARCHAR(50)  = NULL,
    @GroupID           INT          = 0,
    @OrderNo           VARCHAR(50)  = NULL,
    @ProcessID         INT          = 0,
    @VendID            INT          = 0,
    @ItemCode          VARCHAR(50)  = NULL,
    @ShowRepairLots    BIT          = 1,
    @ShowRegularLots   BIT          = 1,
    @ShowLotsOnly      BIT          = 1,
    @ShowNotYetStarted BIT          = 1,
    @Hub               VARCHAR(100) = NULL
)
AS
BEGIN
    SET NOCOUNT ON;

    -- 1. Materialize VRunningLots_Simple ONCE per session (multi-user safe)
    SELECT
        VRL.ItemCode,
        VRL.LotNo,
        VRL.OrderNo,
        VRL.Description,
        VRL.Qty,
        VRL.ProcessID,
        VRL.VendID,
        VRL.ReWorkLot,
        VRL.PolishingItemSNo,
        VRL.SNO,
        VRL.DT,
        QC.TrayNo
    INTO #RunningLots
    FROM dbo.VRunningLots_Simple AS VRL WITH (NOLOCK)
    LEFT OUTER JOIN dbo.VendRcvdDetail_QC_Report AS QC WITH (NOLOCK) ON VRL.VRD_EntryID = QC.VRD_RefID;

    CREATE NONCLUSTERED INDEX IX_TmpRL_LotNo ON #RunningLots (LotNo);
    CREATE NONCLUSTERED INDEX IX_TmpRL_OrderItem ON #RunningLots (OrderNo, ItemCode) INCLUDE (Qty);

    SELECT OrderNo, ItemCode, SUM(Qty) AS RLQty
    INTO #RunningSums
    FROM #RunningLots
    GROUP BY OrderNo, ItemCode;

    CREATE CLUSTERED INDEX IX_TmpRS_OrderItem ON #RunningSums (OrderNo, ItemCode);

    -- 2. Materialize VUnshippedOrderItemList ONCE
    SELECT
        OrderNo,
        CompItemCode,
        DT,
        OrderQty,
        ShippedQty,
        Remarks,
        Price,
        PriceForCost
    INTO #Unshipped
    FROM dbo.VUnshippedOrderItemList WITH (NOLOCK);

    CREATE NONCLUSTERED INDEX IX_TmpUnsh_OrderItem ON #Unshipped (OrderNo, CompItemCode);

    SELECT CompItemCode, SUM(OrderQty - ShippedQty) AS TLUnShippedQty
    INTO #UnshippedByItem
    FROM #Unshipped
    GROUP BY CompItemCode;

    CREATE CLUSTERED INDEX IX_TmpUnshItem ON #UnshippedByItem (CompItemCode);

    -- 3. Combine running lots and not-yet-started items, applying filters
    ;WITH BaseRows AS (
        SELECT
            ISNULL(T1.StartDT, T2.StartDT) AS StartDT,
            RL.LotNo,
            RL.ItemCode,
            RL.Qty,
            RL.Description,
            DATEDIFF(day, RL.DT, GETDATE()) AS LastProcessDays,
            RL.OrderNo,
            CAST('' AS varchar(255)) AS Remarks,
            RL.DT,
            RL.ProcessID,
            RL.ReWorkLot,
            0 AS EntryType,
            pr.Remarks AS POIRemarks,
            0 AS OpenPOQty,
            RL.VendID,
            CASE WHEN RL.SNO <= RL.PolishingItemSNo THEN 1 ELSE 0 END AS R1,
            mpo.MasterPONo,
            RL.TrayNo
        FROM #RunningLots AS RL
        LEFT OUTER JOIN dbo.POI_Remarks AS pr ON RL.LotNo = pr.LotNo
        LEFT OUTER JOIN dbo.VLotsWithMasterPONo AS mpo ON RL.LotNo = mpo.LotNo
        OUTER APPLY (
            SELECT MIN(vr.DT) AS StartDT
            FROM dbo.vendRcvdDetail AS vrd
            INNER JOIN dbo.VendReceived AS vr ON vr.EntryID = vrd.RefID
            WHERE vrd.LotNo = RL.LotNo
        ) AS T1
        OUTER APPLY (
            SELECT MIN(vi.DT) AS StartDT
            FROM dbo.VendIssdDetail AS vid
            INNER JOIN dbo.VendIssued AS vi ON vi.EntryID = vid.RefID
            WHERE vid.LotNo = RL.LotNo
        ) AS T2

        UNION ALL

        SELECT
            u.DT AS StartDT,
            '' AS LotNo,
            u.CompItemCode AS ItemCode,
            u.OrderQty - u.ShippedQty - ISNULL(rs.RLQty, 0) AS Qty,
            '' AS Description,
            0 AS LastProcessDays,
            u.OrderNo,
            u.Remarks,
            u.DT,
            0 AS ProcessID,
            CAST(0 AS bit) AS ReWorkLot,
            1 AS EntryType,
            pr1.Remarks AS POIRemarks,
            ISNULL(TOpenPOs.OpenPOQty, 0) AS OpenPOQty,
            0 AS VendID,
            0 AS R1,
            TOpenPOs.MasterPONo,
            '' AS TrayNo
        FROM #Unshipped AS u
        LEFT OUTER JOIN dbo.POI_Remarks AS pr1
            ON u.OrderNo = pr1.OrderNo
           AND u.CompItemCode = pr1.ItemCode
        LEFT OUTER JOIN #RunningSums AS rs
            ON u.OrderNo = rs.OrderNo
           AND u.CompItemCode = rs.ItemCode
        OUTER APPLY (
            SELECT SUM(op.Qty) AS OpenPOQty, MAX(op.MasterPONo) AS MasterPONo
            FROM dbo.VItemOpenPOs AS op
            WHERE op.OrderNo = u.OrderNo
              AND op.ItemCode = u.CompItemCode
        ) AS TOpenPOs
    ),
    Filtered AS (
        SELECT
            b.StartDT,
            b.LotNo,
            b.ItemCode,
            b.Qty,
            b.Description,
            b.LastProcessDays,
            b.OrderNo,
            b.Remarks,
            b.DT,
            b.ProcessID,
            b.ReWorkLot,
            i.GroupID,
            fco.CustCode,
            fco.Country,
            ISNULL(rs2.RLQty, 0) AS TLRunning,
            ISNULL(rs2.RLQty, 0) AS TLRunning_Order,
            i.ItemName,
            u1.OrderNo AS UnShipped_OrderNo,
            b.EntryType,
            b.POIRemarks,
            b.OpenPOQty,
            b.VendID,
            b.R1 AS RunningLots_R1,
            ISNULL(ubi.TLUnShippedQty, 0) AS TLUnShippedQty,
            u1.Price,
            u1.PriceForCost,
            foi.Price AS FOI_Price,
            fc.Curr,
            b.MasterPONo,
            foi.DeliveryDT AS ItemDeliveryDT,
            pgp.Hub_Name,
            b.TrayNo AS Lot_TrayNo,
            ISNULL(TDispatch.Qty, 0) AS DispatchQty,
            foi.OrderQty
        FROM BaseRows AS b
        INNER JOIN dbo.Items AS i ON b.ItemCode = i.ItemID
        INNER JOIN (
            SELECT OrderNo, CompItemCode, MAX(Price) AS Price, MAX(DeliveryDT) AS DeliveryDT, SUM(Qty) AS OrderQty
            FROM dbo.FOrderItems
            GROUP BY OrderNo, CompItemCode
        ) AS foi ON b.ItemCode = foi.CompItemCode AND b.OrderNo = foi.OrderNo
        LEFT OUTER JOIN dbo.FCustomerOrders AS fco ON fco.OrderNo = b.OrderNo
        LEFT OUTER JOIN dbo.ForeignCustomers AS fc ON fc.CustCode = fco.CustCode AND fc.Country = fco.Country
        LEFT OUTER JOIN #Unshipped AS u1 ON b.ItemCode = u1.CompItemCode AND b.OrderNo = u1.OrderNo
        LEFT OUTER JOIN #RunningSums AS rs2 ON b.OrderNo = rs2.OrderNo AND b.ItemCode = rs2.ItemCode
        LEFT OUTER JOIN #UnshippedByItem AS ubi ON i.ItemID = ubi.CompItemCode
        LEFT OUTER JOIN dbo.ItemProcessGroups AS ipg ON b.ItemCode = ipg.ItemID
        LEFT OUTER JOIN dbo.ProcessGroupsProcesses AS pgp
            ON ipg.PG_RefID = pgp.Group_RefID
           AND b.ProcessID = pgp.Process_RefID
        OUTER APPLY (
            SELECT SUM(dl.InnerQty) AS Qty
            FROM dbo.VDispatchListDetail_Ex_Orderwise AS dl
            WHERE dl.OrderNo = b.OrderNo
              AND dl.ItemCode = b.ItemCode
        ) AS TDispatch
        WHERE (u1.OrderNo IS NOT NULL OR b.OrderNo IN ('Stock-Order', 'Stock-OrderIss'))
          AND (
              (NULLIF(LTRIM(RTRIM(@LotNo)), '') IS NOT NULL AND b.LotNo = LTRIM(RTRIM(@LotNo)))
              OR (
                  NULLIF(LTRIM(RTRIM(@LotNo)), '') IS NULL
                  AND b.Qty > 0
                  AND (@DateRangeMode = 0 OR (b.DT >= @DtFrom AND b.DT <= @DtTo))
                  AND (NULLIF(@CustCode, '0') IS NULL OR fco.CustCode = @CustCode)
                  AND (NULLIF(@Country, '0') IS NULL OR fco.Country = @Country)
                  AND (ISNULL(@GroupID, 0) = 0 OR i.GroupID = @GroupID)
                  AND (NULLIF(@OrderNo, '0') IS NULL OR b.OrderNo = @OrderNo)
                  AND (ISNULL(@ProcessID, 0) = 0 OR b.ProcessID = @ProcessID)
                  AND (ISNULL(@VendID, 0) = 0 OR b.VendID = @VendID)
                  AND (NULLIF(@ItemCode, '0') IS NULL OR b.ItemCode = @ItemCode)
                  AND (@ShowRepairLots = 1 OR b.ReWorkLot = 0)
                  AND (@ShowRegularLots = 1 OR b.ReWorkLot = 1)
                  AND (NULLIF(@Hub, '0') IS NULL OR pgp.Hub_Name = @Hub)
                  AND (
                      (@ShowLotsOnly = 1 AND @ShowNotYetStarted = 1 AND ((b.EntryType <> 0 AND ISNULL(TDispatch.Qty, 0) < foi.OrderQty) OR b.EntryType <> 1))
                      OR (
                          NOT (@ShowLotsOnly = 1 AND @ShowNotYetStarted = 1)
                          AND (@ShowLotsOnly = 1 OR (b.EntryType <> 0 AND ISNULL(TDispatch.Qty, 0) < foi.OrderQty))
                          AND (@ShowNotYetStarted = 1 OR b.EntryType <> 1)
                      )
                  )
              )
          )
    )
    SELECT
        f.StartDT,
        f.LotNo,
        f.ItemCode,
        f.Qty,
        f.Description,
        f.LastProcessDays,
        f.OrderNo,
        f.Remarks,
        f.DT,
        f.ProcessID,
        f.ReWorkLot,
        f.GroupID,
        f.CustCode,
        f.Country,
        f.TLRunning,
        f.TLRunning_Order,
        TSF.R1,
        TSF.R2,
        TPack.PackStock,
        f.ItemName,
        f.UnShipped_OrderNo,
        f.EntryType,
        f.POIRemarks,
        f.OpenPOQty,
        f.VendID,
        f.RunningLots_R1,
        f.TLUnShippedQty,
        TLot_Wastage.Lot_Wastage_Qty,
        f.Price,
        f.PriceForCost,
        f.FOI_Price,
        f.Curr,
        TTray.TrayNo,
        f.MasterPONo,
        f.ItemDeliveryDT,
        f.Hub_Name,
        f.Lot_TrayNo,
        f.DispatchQty,
        f.OrderQty
    FROM Filtered AS f
    OUTER APPLY (
        SELECT
            SUM(CASE WHEN vso.R1 = 1 THEN vso.Qty - vso.IssdQty ELSE 0 END) AS R1,
            SUM(CASE WHEN vso.R1 = 0 THEN vso.Qty - vso.IssdQty ELSE 0 END) AS R2
        FROM dbo.VStockOrderOpening AS vso
        WHERE vso.Qty > vso.IssdQty
          AND vso.ItemID = f.ItemCode
    ) AS TSF
    OUTER APPLY (
        SELECT SUM(ws.NetQty) AS PackStock
        FROM dbo.VItemsWithShelfWiseStock AS ws
        WHERE ws.ItemID = f.ItemCode
    ) AS TPack
    OUTER APPLY (
        SELECT SUM(vrd1.Wastage) AS Lot_Wastage_Qty
        FROM dbo.vendRcvdDetail AS vrd1
        WHERE f.LotNo <> ''
          AND vrd1.LotNo = f.LotNo
    ) AS TLot_Wastage
    OUTER APPLY (
        SELECT TOP (1) qc2.TrayNo
        FROM dbo.VendRcvdDetail_QC_Report AS qc2 WITH (NOLOCK)
        INNER JOIN dbo.vendRcvdDetail AS vrd2 ON vrd2.EntryID = qc2.VRD_RefID
        WHERE f.LotNo <> ''
          AND vrd2.LotNo = f.LotNo
    ) AS TTray
    ORDER BY f.StartDT DESC;
END
GO