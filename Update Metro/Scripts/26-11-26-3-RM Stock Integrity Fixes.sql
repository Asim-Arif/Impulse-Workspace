-- =====================================================================
-- RM Stock Integrity Fixes: VStockRcvd, VStockLedger, VMaterialsForPlacement,
-- UpdateMaterialLocation, and UpdateMaterialLocation_Deletion_SP
-- =====================================================================

CREATE OR ALTER VIEW [dbo].[VStockRcvd]
AS
SELECT
    dbo.VendRcvd.RcvID,
    dbo.VendRcvd.RcvDate,
    dbo.VendRcvdDetailPO.MaterialID,
    dbo.VendRcvd.Description,
    dbo.VVendersAndMakers.AccTitle AS Name,
    ISNULL(dbo.VendRcvdDetailPO.QtyRcvd, 0) - ISNULL(dbo.VendRcvdDetailPO.QtyReturned, 0) + ISNULL(dbo.VendRcvdDetailPO.ExtraQtyRcvd, 0) AS NetQtyRcvd,
    1 AS TransactionType,
    dbo.VendRcvdDetailPO.Rate,
    '' AS StoreRcvdUserName,
    '' AS GateRcvdUserName,
    ISNULL(dbo.VendOrders.IssuanceType, 0) AS IssuanceType,
    ISNULL(dbo.VendOrders.OrderNo, dbo.VendRcvdDetailPO.PORefNo) AS OrderNo,
    dbo.VendRcvdDetailPO.Sheets
FROM dbo.VendRcvdDetailPO
INNER JOIN dbo.VendRcvd
    ON dbo.VendRcvdDetailPO.RcvID = dbo.VendRcvd.RcvID
LEFT OUTER JOIN dbo.VendOrders
    ON dbo.VendRcvdDetailPO.PORefNo = dbo.VendOrders.OrderNo
LEFT OUTER JOIN dbo.VVendersAndMakers
    ON dbo.VendRcvd.VendID = dbo.VVendersAndMakers.AccNo
WHERE (dbo.VendRcvdDetailPO.Checked = 1);
GO

CREATE OR ALTER VIEW [dbo].[VStockLedger]
AS
SELECT
    TransactionID,
    TransactionDate,
    MaterialID,
    Description,
    Reference,
    Quantity,
    TransactionType,
    Rate,
    UserName,
    RecieptID,
    IssuanceType,
    POIssuance,
    Sheets
FROM (
    SELECT
        IssNo AS TransactionID,
        DT AS TransactionDate,
        RMID1 AS MaterialID,
        'Issued to ' + ISNULL(IssdToEmpID, '') + CASE WHEN ISNULL(Remarks, '') <> '' THEN ' ' + Remarks ELSE '' END AS Description,
        CASE
            WHEN ISNULL(VO_OrderNo, '') <> '' THEN AccTitle
            ELSE CASE WHEN ISNULL(IssdToEmpID, '') = '' THEN VenderName ELSE EmpName END
        END AS Reference,
        -Qty AS Quantity,
        0 AS TransactionType,
        0 AS Rate,
        UserName,
        CASE WHEN VI_RefID = 0 THEN VO_OrderNo ELSE RecieptID END AS RecieptID,
        CASE WHEN ForgingGroupID IS NOT NULL THEN CASE WHEN RequiresMaleGrinding = 1 THEN 3 ELSE 2 END ELSE 0 END AS IssuanceType,
        0 AS POIssuance,
        Sheets
    FROM dbo.VMaterialIssued

    UNION ALL

    SELECT
        RcvID,
        RcvDate AS TransactionDate,
        MaterialID,
        Description,
        Name,
        NetQtyRcvd,
        TransactionType,
        Rate,
        StoreRcvdUserName AS UserName,
        OrderNo AS RecieptID,
        IssuanceType,
        0 AS POIssuance,
        Sheets
    FROM dbo.VStockRcvd

    UNION ALL

    SELECT
        OrderNo,
        DT AS TransactionDate,
        RMID1 AS MaterialID,
        'PO Issuance @ ' + CASE IssuanceType WHEN 1 THEN 'Annealing' WHEN 2 THEN 'Machining' WHEN 3 THEN 'Male Grinding' END AS Expr1,
        AccTitle,
        -Qty AS Expr2,
        0 AS TransactionType,
        Rate,
        '' AS UserName,
        OrderNo AS RecieptID,
        IssuanceType,
        1 AS POIssuance,
        0 AS Sheets
    FROM dbo.VVendOrders
    WHERE (IssuanceType <> 0)

    UNION ALL

    SELECT
        OrderNo,
        DT AS TransactionDate,
        MaterialID,
        'Rejection ' AS Description,
        AccTitle,
        -QtyRejected AS Qty,
        0 AS TransactionType,
        Rate,
        UserName,
        RcvID,
        0 AS IssuanceType,
        0 AS POIssuance,
        SheetsRejected
    FROM dbo.VVendRcvdDetailPO_Rejection_Detail
) AS StockLedger;
GO

CREATE OR ALTER VIEW [dbo].[VMaterialsForPlacement]
AS
SELECT
    dbo.VendRcvd.VendID,
    dbo.VendRcvdDetailPO.RcvID,
    dbo.VendRcvdDetailPO.MaterialID,
    dbo.VRMAndItemsUnioned.RMName,
    dbo.VRMAndItemsUnioned.Unit,
    dbo.VendRcvdDetailPO.EntryID,
    ISNULL(TabPlaced.QtyPlaced, 0) AS QtyPlaced,
    dbo.VVendersAndMakers.AccTitle,
    dbo.VendRcvd.RcvDate,
    ISNULL(dbo.VendRcvdDetailPO.QtyRcvd, 0) - ISNULL(TabRej.QtyRejected, 0) + ISNULL(dbo.VendRcvdDetailPO.ExtraQtyRcvd, 0) AS QtyPassed,
    dbo.VendOrderDetail.Qty * dbo.VendOrderDetail.Rate AS Rate,
    dbo.VendRcvdDetailPO.PORefNo,
    dbo.VRMAndItemsUnioned.RMID1,
    dbo.VendRcvdDetailPO.Sheets - ISNULL(TabRej.SheetsRejected, 0) AS Sheets,
    ISNULL(TabPlaced.SheetsPlaced, 0) AS SheetsPlaced,
    TabRej.QtyRejected,
    TabRej.SheetsRejected,
    dbo.VRMAndItemsUnioned.Sample,
    dbo.VendOrderDetail.CCItem,
    dbo.RcvItemsSimpleDetail.EntryID AS RISD_EntryID
FROM dbo.VendRcvd
INNER JOIN dbo.VendRcvdDetailPO
    ON dbo.VendRcvd.RcvID = dbo.VendRcvdDetailPO.RcvID
INNER JOIN dbo.VRMAndItemsUnioned
    ON dbo.VendRcvdDetailPO.MaterialID = dbo.VRMAndItemsUnioned.RMID1
INNER JOIN dbo.VendOrderDetail
    ON dbo.VendRcvdDetailPO.PORefNo = dbo.VendOrderDetail.OrderNo
   AND dbo.VendRcvdDetailPO.MaterialID = dbo.VendOrderDetail.RMID
   AND dbo.VRMAndItemsUnioned.CCItem = dbo.VendOrderDetail.CCItem
INNER JOIN dbo.VendOrders
    ON dbo.VendOrderDetail.OrderNo = dbo.VendOrders.OrderNo
INNER JOIN dbo.VVendersAndMakers
    ON dbo.VendRcvd.VendID = dbo.VVendersAndMakers.AccNo
LEFT OUTER JOIN dbo.RcvItemsSimpleDetail
    ON dbo.VendRcvdDetailPO.EntryID = dbo.RcvItemsSimpleDetail.VRDPO_RefID
LEFT OUTER JOIN (
    SELECT
        m.Rcvd_RefID,
        SUM(ISNULL(m.QtyPlaced, 0)) AS QtyPlaced,
        SUM(ISNULL(m.SheetsPlaced, 0)) AS SheetsPlaced
    FROM dbo.MaterialLocationwiseStatus m
    WHERE NOT EXISTS (
        SELECT 1
        FROM dbo.MaterialMovements mm
        WHERE mm.MLS_To_RefID = m.EntryID
    )
    GROUP BY m.Rcvd_RefID
) AS TabPlaced
    ON dbo.VendRcvdDetailPO.EntryID = TabPlaced.Rcvd_RefID
LEFT OUTER JOIN (
    SELECT
        VRDPO_RefID,
        SUM(QtyRejected) AS QtyRejected,
        SUM(SheetsRejected) AS SheetsRejected
    FROM dbo.VendRcvdDetailPO_Rejection_Detail
    GROUP BY VRDPO_RefID
) AS TabRej
    ON dbo.VendRcvdDetailPO.EntryID = TabRej.VRDPO_RefID
WHERE (dbo.VendRcvd.RcvID NOT IN (SELECT RcvID FROM dbo.RcvIDsToIgnoreForPlacement));
GO

CREATE OR ALTER PROCEDURE [dbo].[UpdateMaterialLocation]
(
    @MaterialID    VARCHAR(50),
    @LocationID    INT,
    @Qty           REAL,
    @Access        BIT         = 0,
    @NewLocationID INT         = 0,
    @UserName      VARCHAR(50) = '',
    @MachineName   VARCHAR(50) = '',
    @Sheets        INT         = 0,
    @PORefNo       VARCHAR(50) = '',
    @DetailEntryID INT         = 0,
    @BatchNo       VARCHAR(50) = '',
    @LotNo         VARCHAR(50) = ''
)
AS
BEGIN
    SET NOCOUNT ON;

    SET @BatchNo       = ISNULL(@BatchNo, '');
    SET @LotNo         = ISNULL(@LotNo, '');
    SET @PORefNo       = ISNULL(@PORefNo, '');
    SET @Sheets        = ISNULL(@Sheets, 0);
    SET @Access        = ISNULL(@Access, 0);
    SET @NewLocationID = ISNULL(@NewLocationID, 0);
    SET @DetailEntryID = ISNULL(@DetailEntryID, 0);

    DECLARE @TargetQty DECIMAL(18, 4) = ROUND(CAST(ISNULL(@Qty, 0) AS DECIMAL(18, 4)), 4);
    DECLARE @TempQty   DECIMAL(18, 4) = 0;
    DECLARE @TempTot   DECIMAL(18, 4) = 0;
    DECLARE @EntryID   INT;
    DECLARE @CurrQty   DECIMAL(18, 4);
    DECLARE @Rcvd_RefID INT;
    DECLARE @RowBatchNo VARCHAR(50);
    DECLARE @RowLotNo   VARCHAR(50);
    DECLARE @RefID     INT;

    IF @TargetQty > 0
    BEGIN
        DECLARE @Cursor_UpdateLocation CURSOR;

        IF @BatchNo = '' AND @LotNo = ''
            SET @Cursor_UpdateLocation = CURSOR LOCAL FAST_FORWARD FOR
                SELECT
                    mls.EntryID,
                    ROUND(CAST(ISNULL(mls.QtyPlaced, 0) AS DECIMAL(18, 4)) - CAST(ISNULL(mls.QtyIssued, 0) AS DECIMAL(18, 4)), 4) AS AvailQty,
                    mls.Rcvd_RefID,
                    ISNULL(mls.BatchNo, ''),
                    ISNULL(mls.LotNo, '')
                FROM dbo.MaterialLocationwiseStatus AS mls WITH (UPDLOCK, ROWLOCK)
                INNER JOIN dbo.VendRcvdDetailPO AS vrdp ON mls.Rcvd_RefID = vrdp.EntryID
                WHERE vrdp.MaterialID = @MaterialID
                  AND mls.Shelf_RefID = @LocationID
                  AND (CAST(ISNULL(mls.QtyPlaced, 0) AS DECIMAL(18, 4)) - CAST(ISNULL(mls.QtyIssued, 0) AS DECIMAL(18, 4))) > 0.0001
                ORDER BY mls.EntryID ASC;
        ELSE
            SET @Cursor_UpdateLocation = CURSOR LOCAL FAST_FORWARD FOR
                SELECT
                    mls.EntryID,
                    ROUND(CAST(ISNULL(mls.QtyPlaced, 0) AS DECIMAL(18, 4)) - CAST(ISNULL(mls.QtyIssued, 0) AS DECIMAL(18, 4)), 4) AS AvailQty,
                    mls.Rcvd_RefID,
                    ISNULL(mls.BatchNo, ''),
                    ISNULL(mls.LotNo, '')
                FROM dbo.MaterialLocationwiseStatus AS mls WITH (UPDLOCK, ROWLOCK)
                INNER JOIN dbo.VendRcvdDetailPO AS vrdp ON mls.Rcvd_RefID = vrdp.EntryID
                WHERE vrdp.MaterialID = @MaterialID
                  AND mls.Shelf_RefID = @LocationID
                  AND ISNULL(mls.BatchNo, '') = @BatchNo
                  AND ISNULL(mls.LotNo, '') = @LotNo
                  AND (CAST(ISNULL(mls.QtyPlaced, 0) AS DECIMAL(18, 4)) - CAST(ISNULL(mls.QtyIssued, 0) AS DECIMAL(18, 4))) > 0.0001
                ORDER BY mls.EntryID ASC;

        OPEN @Cursor_UpdateLocation;
        FETCH NEXT FROM @Cursor_UpdateLocation INTO @EntryID, @CurrQty, @Rcvd_RefID, @RowBatchNo, @RowLotNo;

        WHILE @@FETCH_STATUS = 0 AND @TempTot < @TargetQty
        BEGIN
            IF (@CurrQty > (@TargetQty - @TempTot))
                SET @TempQty = @TargetQty - @TempTot;
            ELSE
                SET @TempQty = @CurrQty;

            IF @TempQty <= 0
                BREAK;

            UPDATE dbo.MaterialLocationwiseStatus
            SET QtyIssued = CAST(ROUND(CAST(ISNULL(QtyIssued, 0) AS DECIMAL(18, 4)) + @TempQty, 4) AS REAL)
            WHERE EntryID = @EntryID;

            IF @DetailEntryID > 0
                INSERT INTO dbo.RMID_MLS_Details (RMID_RefID, MLS_RefID, Qty)
                VALUES (@DetailEntryID, @EntryID, CAST(@TempQty AS REAL));

            IF @Access = 1
            BEGIN
                INSERT INTO dbo.MaterialLocationwiseStatus (Rcvd_RefID, Shelf_RefID, QtyPlaced, QtyIssued, UserName, MachineName, BatchNo, LotNo)
                VALUES (@Rcvd_RefID, @NewLocationID, CAST(@TempQty AS REAL), 0, @UserName, @MachineName, @RowBatchNo, @RowLotNo);

                SET @RefID = CAST(SCOPE_IDENTITY() AS INT);

                INSERT INTO dbo.MaterialMovements (MLS_From_RefID, MLS_To_RefID, UserName, MachineName, QtyMoved)
                VALUES (@EntryID, @RefID, @UserName, @MachineName, CAST(@TempQty AS REAL));
            END

            SET @TempTot = @TempTot + @TempQty;
            FETCH NEXT FROM @Cursor_UpdateLocation INTO @EntryID, @CurrQty, @Rcvd_RefID, @RowBatchNo, @RowLotNo;
        END

        CLOSE @Cursor_UpdateLocation;
        DEALLOCATE @Cursor_UpdateLocation;

        IF (@TargetQty - @TempTot) > 0.001
        BEGIN
            DECLARE @ErrMsg NVARCHAR(400) =
                'Insufficient shelf stock for material ' + ISNULL(@MaterialID, '') +
                ' on location #' + CAST(ISNULL(@LocationID, 0) AS VARCHAR(20)) +
                '. Requested: ' + CAST(@TargetQty AS VARCHAR(30)) +
                ', Available: ' + CAST(@TempTot AS VARCHAR(30)) + '.';
            THROW 51001, @ErrMsg, 1;
        END
    END

    IF @Sheets > 0
    BEGIN
        DECLARE @Cursor_UpdateLocationSheets CURSOR;
        DECLARE @CurrSheets INT;
        DECLARE @TempSheets INT = 0;
        DECLARE @TempSheetsTot INT = 0;

        IF @PORefNo = ''
            SET @Cursor_UpdateLocationSheets = CURSOR LOCAL FAST_FORWARD FOR
                SELECT
                    mls.EntryID,
                    ISNULL(mls.SheetsPlaced, 0) - ISNULL(mls.SheetsIssued, 0),
                    mls.Rcvd_RefID,
                    ISNULL(mls.BatchNo, ''),
                    ISNULL(mls.LotNo, '')
                FROM dbo.MaterialLocationwiseStatus AS mls WITH (UPDLOCK, ROWLOCK)
                INNER JOIN dbo.VendRcvdDetailPO AS vrdp ON mls.Rcvd_RefID = vrdp.EntryID
                WHERE (ISNULL(mls.SheetsPlaced, 0) - ISNULL(mls.SheetsIssued, 0)) > 0
                  AND vrdp.MaterialID = @MaterialID
                  AND mls.Shelf_RefID = @LocationID
                ORDER BY mls.EntryID ASC;
        ELSE
            SET @Cursor_UpdateLocationSheets = CURSOR LOCAL FAST_FORWARD FOR
                SELECT
                    mls.EntryID,
                    ISNULL(mls.SheetsPlaced, 0) - ISNULL(mls.SheetsIssued, 0),
                    mls.Rcvd_RefID,
                    ISNULL(mls.BatchNo, ''),
                    ISNULL(mls.LotNo, '')
                FROM dbo.MaterialLocationwiseStatus AS mls WITH (UPDLOCK, ROWLOCK)
                INNER JOIN dbo.VendRcvdDetailPO AS vrdp ON mls.Rcvd_RefID = vrdp.EntryID
                WHERE (ISNULL(mls.SheetsPlaced, 0) - ISNULL(mls.SheetsIssued, 0)) > 0
                  AND vrdp.MaterialID = @MaterialID
                  AND mls.Shelf_RefID = @LocationID
                  AND vrdp.PORefNo = @PORefNo
                ORDER BY mls.EntryID ASC;

        OPEN @Cursor_UpdateLocationSheets;
        FETCH NEXT FROM @Cursor_UpdateLocationSheets INTO @EntryID, @CurrSheets, @Rcvd_RefID, @RowBatchNo, @RowLotNo;

        WHILE @@FETCH_STATUS = 0 AND @TempSheetsTot < @Sheets
        BEGIN
            IF (@CurrSheets > (@Sheets - @TempSheetsTot))
                SET @TempSheets = @Sheets - @TempSheetsTot;
            ELSE
                SET @TempSheets = @CurrSheets;

            IF @TempSheets <= 0
                BREAK;

            UPDATE dbo.MaterialLocationwiseStatus
            SET SheetsIssued = ISNULL(SheetsIssued, 0) + @TempSheets
            WHERE EntryID = @EntryID;

            IF @DetailEntryID > 0
                INSERT INTO dbo.RMID_MLS_Sheets_Details (RMID_RefID, MLS_RefID, Sheets)
                VALUES (@DetailEntryID, @EntryID, @TempSheets);

            IF @Access = 1
            BEGIN
                INSERT INTO dbo.MaterialLocationwiseStatus (Rcvd_RefID, Shelf_RefID, SheetsPlaced, SheetsIssued, UserName, MachineName, BatchNo, LotNo)
                VALUES (@Rcvd_RefID, @NewLocationID, @TempSheets, 0, @UserName, @MachineName, @RowBatchNo, @RowLotNo);

                SET @RefID = CAST(SCOPE_IDENTITY() AS INT);

                INSERT INTO dbo.MaterialMovements (MLS_From_RefID, MLS_To_RefID, UserName, MachineName)
                VALUES (@EntryID, @RefID, @UserName, @MachineName);
            END

            SET @TempSheetsTot = @TempSheetsTot + @TempSheets;
            FETCH NEXT FROM @Cursor_UpdateLocationSheets INTO @EntryID, @CurrSheets, @Rcvd_RefID, @RowBatchNo, @RowLotNo;
        END

        CLOSE @Cursor_UpdateLocationSheets;
        DEALLOCATE @Cursor_UpdateLocationSheets;
    END
END
GO

CREATE OR ALTER PROCEDURE [dbo].[UpdateMaterialLocation_Deletion_SP]
(
    @IssNo VARCHAR(255)
)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @StartedTran BIT = 0;
    IF @@TRANCOUNT = 0
    BEGIN
        BEGIN TRAN;
        SET @StartedTran = 1;
    END

    BEGIN TRY
        -- 1. Restore MaterialLocationwiseStatus.QtyIssued (aggregated per MLS row)
        UPDATE mls
        SET mls.QtyIssued = CASE
            WHEN ROUND(CAST(ISNULL(mls.QtyIssued, 0) AS DECIMAL(18, 4)) - Agg.TotalQty, 4) < 0 THEN 0
            ELSE CAST(ROUND(CAST(ISNULL(mls.QtyIssued, 0) AS DECIMAL(18, 4)) - Agg.TotalQty, 4) AS REAL)
        END
        FROM dbo.MaterialLocationwiseStatus AS mls
        INNER JOIN (
            SELECT d.MLS_RefID, SUM(CAST(ISNULL(d.Qty, 0) AS DECIMAL(18, 4))) AS TotalQty
            FROM dbo.RMID_MLS_Details AS d
            INNER JOIN dbo.RawMaterialIssuanceDetail AS rmid ON d.RMID_RefID = rmid.EntryID
            WHERE rmid.IssNo = @IssNo
            GROUP BY d.MLS_RefID
        ) AS Agg ON mls.EntryID = Agg.MLS_RefID;

        -- 2. Restore MaterialLocationwiseStatus.SheetsIssued (aggregated per MLS row)
        UPDATE mls
        SET mls.SheetsIssued = CASE
            WHEN ISNULL(mls.SheetsIssued, 0) - Agg.TotalSheets < 0 THEN 0
            ELSE ISNULL(mls.SheetsIssued, 0) - Agg.TotalSheets
        END
        FROM dbo.MaterialLocationwiseStatus AS mls
        INNER JOIN (
            SELECT d.MLS_RefID, SUM(ISNULL(d.Sheets, 0)) AS TotalSheets
            FROM dbo.RMID_MLS_Sheets_Details AS d
            INNER JOIN dbo.RawMaterialIssuanceDetail AS rmid ON d.RMID_RefID = rmid.EntryID
            WHERE rmid.IssNo = @IssNo
            GROUP BY d.MLS_RefID
        ) AS Agg ON mls.EntryID = Agg.MLS_RefID;

        -- 3. Restore RM.QtyInStock (+ Qty, aggregated per RMID1)
        UPDATE rm
        SET rm.QtyInStock = ISNULL(rm.QtyInStock, 0) + Agg.TotalQty
        FROM dbo.RM AS rm
        INNER JOIN (
            SELECT rmid.RMID1, SUM(ISNULL(rmid.Qty, 0)) AS TotalQty
            FROM dbo.RawMaterialIssuanceDetail AS rmid
            WHERE rmid.IssNo = @IssNo
            GROUP BY rmid.RMID1
        ) AS Agg ON rm.RMID1 = Agg.RMID1;

        -- 4. Delete link, detail, and header records
        DELETE FROM dbo.RMID_MLS_Details
        WHERE RMID_RefID IN (SELECT EntryID FROM dbo.RawMaterialIssuanceDetail WHERE IssNo = @IssNo);

        DELETE FROM dbo.RMID_MLS_Sheets_Details
        WHERE RMID_RefID IN (SELECT EntryID FROM dbo.RawMaterialIssuanceDetail WHERE IssNo = @IssNo);

        DELETE FROM dbo.RawMaterialIssuanceDetail WHERE IssNo = @IssNo;
        DELETE FROM dbo.RawMaterialIssuance WHERE IssNo = @IssNo;

        IF @StartedTran = 1
            COMMIT TRAN;
    END TRY
    BEGIN CATCH
        IF @StartedTran = 1 AND @@TRANCOUNT > 0
            ROLLBACK TRAN;
        THROW;
    END CATCH
END
GO
