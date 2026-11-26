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

    CREATE CLUSTERED INDEX IX_TmpUnsh_OrderItem ON #Unshipped (OrderNo, CompItemCode);

    SELECT CompItemCode, SUM(OrderQty - ShippedQty) AS TLUnShippedQty
    INTO #UnshippedByItem
    FROM #Unshipped
    GROUP BY CompItemCode;

    CREATE CLUSTERED INDEX IX_TmpUnshItem ON #UnshippedByItem (CompItemCode);

    -- 3. Materialize small lookup/aggregation views ONCE instead of per-row OUTER APPLY
    SELECT OrderNo, ItemCode, SUM(Qty) AS OpenPOQty, MAX(MasterPONo) AS MasterPONo
    INTO #OpenPOs
    FROM dbo.VItemOpenPOs WITH (NOLOCK)
    GROUP BY OrderNo, ItemCode;

    CREATE CLUSTERED INDEX IX_TmpOpenPOs ON #OpenPOs (OrderNo, ItemCode);

    SELECT OrderNo, ItemCode, SUM(InnerQty) AS Qty
    INTO #Dispatch
    FROM dbo.VDispatchListDetail_Ex_Orderwise WITH (NOLOCK)
    GROUP BY OrderNo, ItemCode;

    CREATE CLUSTERED INDEX IX_TmpDispatch ON #Dispatch (OrderNo, ItemCode);

    SELECT
        ItemID,
        SUM(CASE WHEN R1 = 1 THEN Qty - IssdQty ELSE 0 END) AS R1,
        SUM(CASE WHEN R1 = 0 THEN Qty - IssdQty ELSE 0 END) AS R2
    INTO #SF
    FROM dbo.VStockOrderOpening WITH (NOLOCK)
    WHERE Qty > IssdQty
    GROUP BY ItemID;

    CREATE CLUSTERED INDEX IX_TmpSF ON #SF (ItemID);

    SELECT ItemID, SUM(NetQty) AS PackStock
    INTO #PackStock
    FROM dbo.VItemsWithShelfWiseStock WITH (NOLOCK)
    GROUP BY ItemID;

    CREATE CLUSTERED INDEX IX_TmpPackStock ON #PackStock (ItemID);

    -- 4. Compute lot-level dates, wastage, and tray ONLY for active running lots
    SELECT DISTINCT LotNo INTO #ActiveLotNos FROM #RunningLots WHERE LotNo <> '';
    CREATE CLUSTERED INDEX IX_TmpActiveLots ON #ActiveLotNos (LotNo);

    SELECT vrd.LotNo, MIN(vr.DT) AS StartDT
    INTO #LotStartRcvd
    FROM dbo.vendRcvdDetail AS vrd WITH (NOLOCK)
    INNER JOIN #ActiveLotNos AS al ON vrd.LotNo = al.LotNo
    INNER JOIN dbo.VendReceived AS vr WITH (NOLOCK) ON vr.EntryID = vrd.RefID
    GROUP BY vrd.LotNo;

    CREATE CLUSTERED INDEX IX_TmpLotRcvd ON #LotStartRcvd (LotNo);

    SELECT vid.LotNo, MIN(vi.DT) AS StartDT
    INTO #LotStartIssd
    FROM dbo.VendIssdDetail AS vid WITH (NOLOCK)
    INNER JOIN #ActiveLotNos AS al ON vid.LotNo = al.LotNo
    INNER JOIN dbo.VendIssued AS vi WITH (NOLOCK) ON vi.EntryID = vid.RefID
    GROUP BY vid.LotNo;

    CREATE CLUSTERED INDEX IX_TmpLotIssd ON #LotStartIssd (LotNo);

    SELECT vrd.LotNo, SUM(vrd.Wastage) AS Lot_Wastage_Qty
    INTO #LotWastage
    FROM dbo.vendRcvdDetail AS vrd WITH (NOLOCK)
    INNER JOIN #ActiveLotNos AS al ON vrd.LotNo = al.LotNo
    GROUP BY vrd.LotNo;

    CREATE CLUSTERED INDEX IX_TmpLotWastage ON #LotWastage (LotNo);

    SELECT vrd.LotNo, MAX(qc.TrayNo) AS TrayNo
    INTO #LotTray
    FROM dbo.VendRcvdDetail_QC_Report AS qc WITH (NOLOCK)
    INNER JOIN dbo.vendRcvdDetail AS vrd WITH (NOLOCK) ON vrd.EntryID = qc.VRD_RefID
    INNER JOIN #ActiveLotNos AS al ON vrd.LotNo = al.LotNo
    GROUP BY vrd.LotNo;

    CREATE CLUSTERED INDEX IX_TmpLotTray ON #LotTray (LotNo);

    -- 5. Combine running lots and not-yet-started items, applying filters
    ;WITH BaseRows AS (
        SELECT
            ISNULL(lsr.StartDT, lsi.StartDT) AS StartDT,
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
        LEFT OUTER JOIN dbo.POI_Remarks AS pr WITH (NOLOCK) ON RL.LotNo = pr.LotNo
        LEFT OUTER JOIN dbo.VLotsWithMasterPONo AS mpo WITH (NOLOCK) ON RL.LotNo = mpo.LotNo
        LEFT OUTER JOIN #LotStartRcvd AS lsr ON RL.LotNo = lsr.LotNo
        LEFT OUTER JOIN #LotStartIssd AS lsi ON RL.LotNo = lsi.LotNo

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
            ISNULL(opo.OpenPOQty, 0) AS OpenPOQty,
            0 AS VendID,
            0 AS R1,
            opo.MasterPONo,
            '' AS TrayNo
        FROM #Unshipped AS u
        LEFT OUTER JOIN dbo.POI_Remarks AS pr1 WITH (NOLOCK)
            ON u.OrderNo = pr1.OrderNo
           AND u.CompItemCode = pr1.ItemCode
        LEFT OUTER JOIN #RunningSums AS rs
            ON u.OrderNo = rs.OrderNo
           AND u.CompItemCode = rs.ItemCode
        LEFT OUTER JOIN #OpenPOs AS opo
            ON u.OrderNo = opo.OrderNo
           AND u.CompItemCode = opo.ItemCode
    )
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
        tsf.R1,
        tsf.R2,
        tp.PackStock,
        i.ItemName,
        u1.OrderNo AS UnShipped_OrderNo,
        b.EntryType,
        b.POIRemarks,
        b.OpenPOQty,
        b.VendID,
        b.R1 AS RunningLots_R1,
        ISNULL(ubi.TLUnShippedQty, 0) AS TLUnShippedQty,
        lw.Lot_Wastage_Qty,
        u1.Price,
        u1.PriceForCost,
        foi.Price AS FOI_Price,
        fc.Curr,
        lt.TrayNo,
        b.MasterPONo,
        foi.DeliveryDT AS ItemDeliveryDT,
        pgp.Hub_Name,
        b.TrayNo AS Lot_TrayNo,
        ISNULL(td.Qty, 0) AS DispatchQty,
        foi.OrderQty
    FROM BaseRows AS b
    INNER JOIN dbo.Items AS i WITH (NOLOCK) ON b.ItemCode = i.ItemID
    INNER JOIN (
        SELECT OrderNo, CompItemCode, MAX(Price) AS Price, MAX(DeliveryDT) AS DeliveryDT, SUM(Qty) AS OrderQty
        FROM dbo.FOrderItems WITH (NOLOCK)
        GROUP BY OrderNo, CompItemCode
    ) AS foi ON b.ItemCode = foi.CompItemCode AND b.OrderNo = foi.OrderNo
    LEFT OUTER JOIN dbo.FCustomerOrders AS fco WITH (NOLOCK) ON fco.OrderNo = b.OrderNo
    LEFT OUTER JOIN dbo.ForeignCustomers AS fc WITH (NOLOCK) ON fc.CustCode = fco.CustCode AND fc.Country = fco.Country
    LEFT OUTER JOIN #Unshipped AS u1 ON b.ItemCode = u1.CompItemCode AND b.OrderNo = u1.OrderNo
    LEFT OUTER JOIN #RunningSums AS rs2 ON b.OrderNo = rs2.OrderNo AND b.ItemCode = rs2.ItemCode
    LEFT OUTER JOIN #UnshippedByItem AS ubi ON i.ItemID = ubi.CompItemCode
    LEFT OUTER JOIN dbo.ItemProcessGroups AS ipg WITH (NOLOCK) ON b.ItemCode = ipg.ItemID
    LEFT OUTER JOIN dbo.ProcessGroupsProcesses AS pgp WITH (NOLOCK)
        ON ipg.PG_RefID = pgp.Group_RefID
       AND b.ProcessID = pgp.Process_RefID
    LEFT OUTER JOIN #Dispatch AS td ON b.OrderNo = td.OrderNo AND b.ItemCode = td.ItemCode
    LEFT OUTER JOIN #SF AS tsf ON b.ItemCode = tsf.ItemID
    LEFT OUTER JOIN #PackStock AS tp ON b.ItemCode = tp.ItemID
    LEFT OUTER JOIN #LotWastage AS lw ON b.LotNo = lw.LotNo
    LEFT OUTER JOIN #LotTray AS lt ON b.LotNo = lt.LotNo
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
                  (@ShowLotsOnly = 1 AND @ShowNotYetStarted = 1 AND ((b.EntryType <> 0 AND ISNULL(td.Qty, 0) < foi.OrderQty) OR b.EntryType <> 1))
                  OR (
                      NOT (@ShowLotsOnly = 1 AND @ShowNotYetStarted = 1)
                      AND (@ShowLotsOnly = 1 OR (b.EntryType <> 0 AND ISNULL(td.Qty, 0) < foi.OrderQty))
                      AND (@ShowNotYetStarted = 1 OR b.EntryType <> 1)
                  )
              )
          )
      )
    ORDER BY b.StartDT DESC;
END
GO