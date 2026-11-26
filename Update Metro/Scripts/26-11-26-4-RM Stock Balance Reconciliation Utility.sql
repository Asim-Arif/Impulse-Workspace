-- =====================================================================
-- RM Stock Balance Reconciliation Utility (Database-Independent)
-- =====================================================================
-- PURPOSE:
--   Audits and optionally reconciles discrepancies between:
--     1. RM Stock Ledger Balance  = RM.OpeningStock + SUM(VStockLedger.Quantity)
--     2. Shelf / VRM Balance      = SUM(VMaterialLocationANDPOWiseBalances.ShelfQty)
--                                   (same as VRM.QtyInStock)
--
-- IMPORTANT:
--   Run MASTER_MIGRATION_ALL_SCRIPTS.sql FIRST so VStockRcvd, VStockLedger,
--   VMaterialsForPlacement, UpdateMaterialLocation, and
--   UpdateMaterialLocation_Deletion_SP are updated before reconciling.
--   (Fixing VStockRcvd alone resolves all missing-PO ledger drops automatically!)
--
-- MODES (change @Mode below):
--   'AUDIT'                -> Read-only report of all mismatched materials & exact causes (DEFAULT)
--   'SYNC_LEDGER_TO_SHELF' -> Adjusts RM.OpeningStock so RM Stock Ledger matches Shelf/VRM Balance
--                             (use when physical store shelf counts on Issue Material are accurate)
--   'SYNC_SHELF_TO_LEDGER' -> Deducts excess shelf stock from oldest open MaterialLocationwiseStatus
--                             rows (FIFO, clamped at 0) and clears negative ledger balances via
--                             RM.OpeningStock so both Ledger and Shelf agree.
-- =====================================================================

SET NOCOUNT ON;

DECLARE @Mode VARCHAR(30) = 'SYNC_LEDGER_TO_SHELF'; -- 'AUDIT' | 'SYNC_LEDGER_TO_SHELF' | 'SYNC_SHELF_TO_LEDGER'

-- Build snapshot of current balances and root-cause diagnostics
IF OBJECT_ID('tempdb..#RMReconcile') IS NOT NULL DROP TABLE #RMReconcile;

SELECT
    r.RMID,
    r.RMID1,
    r.RMName,
    ROUND(CAST(ISNULL(r.OpeningStock, 0) AS DECIMAL(18, 4)), 4) AS OpeningStock,
    ROUND(CAST(ISNULL(r.OpeningStock, 0) AS DECIMAL(18, 4)) + ISNULL(l.LedgerSum, 0), 4) AS LedgerBal,
    ROUND(ISNULL(s.ShelfSum, 0), 4) AS ShelfBal,
    ROUND((CAST(ISNULL(r.OpeningStock, 0) AS DECIMAL(18, 4)) + ISNULL(l.LedgerSum, 0)) - ISNULL(s.ShelfSum, 0), 4) AS LedgerMinusShelf,
    ROUND(ISNULL(rc.CheckedRcvd, 0), 4) AS CheckedRcvd,
    ROUND(ISNULL(ip.InitialPlaced, 0), 4) AS InitialPlaced,
    ROUND(CASE WHEN ISNULL(rc.CheckedRcvd, 0) - ISNULL(ip.InitialPlaced, 0) > 0.001 THEN ISNULL(rc.CheckedRcvd, 0) - ISNULL(ip.InitialPlaced, 0) ELSE 0 END, 4) AS UnplacedQty,
    ROUND(CASE WHEN ISNULL(ip.InitialPlaced, 0) - ISNULL(rc.CheckedRcvd, 0) > 0.001 THEN ISNULL(ip.InitialPlaced, 0) - ISNULL(rc.CheckedRcvd, 0) ELSE 0 END, 4) AS OverPlacedQty,
    ROUND(ISNULL(iss.TotalRMIDIssued, 0), 4) AS TotalRMIDIssued,
    ROUND(ISNULL(mls_iss.TotalMLSDeducted, 0), 4) AS TotalMLSDeducted,
    ROUND(ISNULL(iss.TotalRMIDIssued, 0) - ISNULL(mls_iss.TotalMLSDeducted, 0), 4) AS UnDeductedIssuanceQty,
    ROUND(ISNULL(orp.OrphanPORcvdQty, 0), 4) AS OrphanPORcvdQty
INTO #RMReconcile
FROM dbo.RM AS r
LEFT JOIN (
    SELECT MaterialID, SUM(CAST(ISNULL(Quantity, 0) AS DECIMAL(18, 4))) AS LedgerSum
    FROM dbo.VStockLedger
    GROUP BY MaterialID
) AS l ON l.MaterialID = r.RMID1
LEFT JOIN (
    SELECT MaterialID, SUM(CAST(ISNULL(ShelfQty, 0) AS DECIMAL(18, 4))) AS ShelfSum
    FROM dbo.VMaterialLocationANDPOWiseBalances
    GROUP BY MaterialID
) AS s ON s.MaterialID = r.RMID1
LEFT JOIN (
    SELECT MaterialID,
           SUM(CAST(ISNULL(QtyRcvd, 0) AS DECIMAL(18, 4)) - CAST(ISNULL(QtyReturned, 0) AS DECIMAL(18, 4)) + CAST(ISNULL(ExtraQtyRcvd, 0) AS DECIMAL(18, 4))) AS CheckedRcvd
    FROM dbo.VendRcvdDetailPO
    WHERE Checked = 1
    GROUP BY MaterialID
) AS rc ON rc.MaterialID = r.RMID1
LEFT JOIN (
    SELECT vrdp.MaterialID,
           SUM(CAST(ISNULL(m.QtyPlaced, 0) AS DECIMAL(18, 4))) AS InitialPlaced
    FROM dbo.MaterialLocationwiseStatus AS m
    INNER JOIN dbo.VendRcvdDetailPO AS vrdp ON m.Rcvd_RefID = vrdp.EntryID
    WHERE NOT EXISTS (SELECT 1 FROM dbo.MaterialMovements AS mm WHERE mm.MLS_To_RefID = m.EntryID)
    GROUP BY vrdp.MaterialID
) AS ip ON ip.MaterialID = r.RMID1
LEFT JOIN (
    SELECT RMID1, SUM(CAST(ISNULL(Qty, 0) AS DECIMAL(18, 4))) AS TotalRMIDIssued
    FROM dbo.RawMaterialIssuanceDetail
    GROUP BY RMID1
) AS iss ON iss.RMID1 = r.RMID1
LEFT JOIN (
    SELECT rmid.RMID1, SUM(CAST(ISNULL(d.Qty, 0) AS DECIMAL(18, 4))) AS TotalMLSDeducted
    FROM dbo.RMID_MLS_Details AS d
    INNER JOIN dbo.RawMaterialIssuanceDetail AS rmid ON d.RMID_RefID = rmid.EntryID
    GROUP BY rmid.RMID1
) AS mls_iss ON mls_iss.RMID1 = r.RMID1
LEFT JOIN (
    SELECT vrdp.MaterialID,
           SUM(CAST(ISNULL(vrdp.QtyRcvd, 0) AS DECIMAL(18, 4)) - CAST(ISNULL(vrdp.QtyReturned, 0) AS DECIMAL(18, 4)) + CAST(ISNULL(vrdp.ExtraQtyRcvd, 0) AS DECIMAL(18, 4))) AS OrphanPORcvdQty
    FROM dbo.VendRcvdDetailPO AS vrdp
    LEFT JOIN dbo.VendOrders AS vo ON vo.OrderNo = vrdp.PORefNo
    WHERE vrdp.Checked = 1 AND vo.OrderNo IS NULL
    GROUP BY vrdp.MaterialID
) AS orp ON orp.MaterialID = r.RMID1;

-- Always display the current audit table
SELECT
    RMID,
    RMID1,
    RMName,
    OpeningStock,
    LedgerBal,
    ShelfBal,
    LedgerMinusShelf,
    UnplacedQty,
    OverPlacedQty,
    UnDeductedIssuanceQty,
    OrphanPORcvdQty
FROM #RMReconcile
WHERE ABS(LedgerMinusShelf) > 0.005
ORDER BY RMID1;

IF @Mode = 'AUDIT'
BEGIN
    PRINT 'AUDIT ONLY — no rows were modified. Set @Mode = ''SYNC_LEDGER_TO_SHELF'' or ''SYNC_SHELF_TO_LEDGER'' to apply reconciliation.';
    DROP TABLE #RMReconcile;
    RETURN;
END

BEGIN TRY
    BEGIN TRAN;

    IF @Mode = 'SYNC_LEDGER_TO_SHELF'
    BEGIN
        -- Adjust RM.OpeningStock so LedgerBal == ShelfBal (VRM.QtyInStock)
        UPDATE r
        SET r.OpeningStock = CAST(ROUND(t.OpeningStock + (t.ShelfBal - t.LedgerBal), 4) AS REAL),
            r.QtyInStock   = CAST(t.ShelfBal AS REAL)
        FROM dbo.RM AS r
        INNER JOIN #RMReconcile AS t ON t.RMID = r.RMID
        WHERE ABS(t.LedgerMinusShelf) > 0.005;

        -- Also keep legacy RM.QtyInStock aligned with VRM/ShelfBal everywhere
        UPDATE r
        SET r.QtyInStock = CAST(t.ShelfBal AS REAL)
        FROM dbo.RM AS r
        INNER JOIN #RMReconcile AS t ON t.RMID = r.RMID
        WHERE ABS(ISNULL(r.QtyInStock, 0) - t.ShelfBal) > 0.005;

        PRINT 'SYNC_LEDGER_TO_SHELF completed successfully.';
    END
    ELSE IF @Mode = 'SYNC_SHELF_TO_LEDGER'
    BEGIN
        DECLARE @CurRMID INT, @CurRMID1 VARCHAR(50), @CurLedger DECIMAL(18,4), @CurShelf DECIMAL(18,4), @CurOpening DECIMAL(18,4);
        DECLARE @TargetShelf DECIMAL(18,4), @ToDeduct DECIMAL(18,4);

        DECLARE cur_items CURSOR LOCAL FAST_FORWARD FOR
            SELECT RMID, RMID1, LedgerBal, ShelfBal, OpeningStock
            FROM #RMReconcile
            WHERE ShelfBal > LedgerBal + 0.005;

        OPEN cur_items;
        FETCH NEXT FROM cur_items INTO @CurRMID, @CurRMID1, @CurLedger, @CurShelf, @CurOpening;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            -- Physical shelves cannot hold negative stock; clamp target at 0
            SET @TargetShelf = CASE WHEN @CurLedger < 0 THEN 0 ELSE @CurLedger END;
            SET @ToDeduct = ROUND(@CurShelf - @TargetShelf, 4);

            -- Deduct excess from oldest open MaterialLocationwiseStatus rows (FIFO)
            DECLARE @MlsEntryID INT, @MlsAvail DECIMAL(18,4), @DeductHere DECIMAL(18,4);
            DECLARE cur_mls CURSOR LOCAL FAST_FORWARD FOR
                SELECT mls.EntryID,
                       ROUND(CAST(ISNULL(mls.QtyPlaced, 0) AS DECIMAL(18, 4)) - CAST(ISNULL(mls.QtyIssued, 0) AS DECIMAL(18, 4)), 4)
                FROM dbo.MaterialLocationwiseStatus AS mls WITH (UPDLOCK, ROWLOCK)
                INNER JOIN dbo.VendRcvdDetailPO AS vrdp ON mls.Rcvd_RefID = vrdp.EntryID
                WHERE vrdp.MaterialID = @CurRMID1
                  AND (CAST(ISNULL(mls.QtyPlaced, 0) AS DECIMAL(18, 4)) - CAST(ISNULL(mls.QtyIssued, 0) AS DECIMAL(18, 4))) > 0.0001
                ORDER BY mls.EntryID ASC;

            OPEN cur_mls;
            FETCH NEXT FROM cur_mls INTO @MlsEntryID, @MlsAvail;

            WHILE @@FETCH_STATUS = 0 AND @ToDeduct > 0.0001
            BEGIN
                SET @DeductHere = CASE WHEN @MlsAvail > @ToDeduct THEN @ToDeduct ELSE @MlsAvail END;

                UPDATE dbo.MaterialLocationwiseStatus
                SET QtyIssued = CAST(ROUND(CAST(ISNULL(QtyIssued, 0) AS DECIMAL(18, 4)) + @DeductHere, 4) AS REAL)
                WHERE EntryID = @MlsEntryID;

                SET @ToDeduct = ROUND(@ToDeduct - @DeductHere, 4);
                FETCH NEXT FROM cur_mls INTO @MlsEntryID, @MlsAvail;
            END

            CLOSE cur_mls;
            DEALLOCATE cur_mls;

            -- If LedgerBal was negative, bring Ledger up to 0 via RM.OpeningStock so Ledger == Shelf == 0
            IF @CurLedger < 0
            BEGIN
                UPDATE dbo.RM
                SET OpeningStock = CAST(ROUND(@CurOpening + (-@CurLedger), 4) AS REAL),
                    QtyInStock = 0
                WHERE RMID = @CurRMID;
            END
            ELSE
            BEGIN
                UPDATE dbo.RM
                SET QtyInStock = CAST(@TargetShelf AS REAL)
                WHERE RMID = @CurRMID;
            END

            FETCH NEXT FROM cur_items INTO @CurRMID, @CurRMID1, @CurLedger, @CurShelf, @CurOpening;
        END

        CLOSE cur_items;
        DEALLOCATE cur_items;

        PRINT 'SYNC_SHELF_TO_LEDGER completed. Note: Any items with LedgerBal > ShelfBal have unplaced receipts that should be placed via Material Placement.';
    END
    ELSE
    BEGIN
        THROW 51002, 'Invalid @Mode. Use AUDIT, SYNC_LEDGER_TO_SHELF, or SYNC_SHELF_TO_LEDGER.', 1;
    END

    COMMIT TRAN;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0 ROLLBACK TRAN;
    THROW;
END CATCH

-- Post-reconciliation verification
SELECT
    r.RMID,
    r.RMID1,
    r.RMName,
    ROUND(CAST(ISNULL(r.OpeningStock, 0) AS DECIMAL(18, 4)), 4) AS OpeningStock,
    ROUND(CAST(ISNULL(r.OpeningStock, 0) AS DECIMAL(18, 4)) + ISNULL(l.LedgerSum, 0), 4) AS PostLedgerBal,
    ROUND(ISNULL(s.ShelfSum, 0), 4) AS PostShelfBal,
    ROUND((CAST(ISNULL(r.OpeningStock, 0) AS DECIMAL(18, 4)) + ISNULL(l.LedgerSum, 0)) - ISNULL(s.ShelfSum, 0), 4) AS RemainingDiff
FROM dbo.RM AS r
LEFT JOIN (
    SELECT MaterialID, SUM(CAST(ISNULL(Quantity, 0) AS DECIMAL(18, 4))) AS LedgerSum
    FROM dbo.VStockLedger
    GROUP BY MaterialID
) AS l ON l.MaterialID = r.RMID1
LEFT JOIN (
    SELECT MaterialID, SUM(CAST(ISNULL(ShelfQty, 0) AS DECIMAL(18, 4))) AS ShelfSum
    FROM dbo.VMaterialLocationANDPOWiseBalances
    GROUP BY MaterialID
) AS s ON s.MaterialID = r.RMID1
WHERE ABS((CAST(ISNULL(r.OpeningStock, 0) AS DECIMAL(18, 4)) + ISNULL(l.LedgerSum, 0)) - ISNULL(s.ShelfSum, 0)) > 0.005
ORDER BY r.RMID1;

DROP TABLE #RMReconcile;
GO
