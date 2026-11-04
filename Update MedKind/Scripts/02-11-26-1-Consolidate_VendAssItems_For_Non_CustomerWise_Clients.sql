-- ==============================================================================
-- Script: Consolidate VendAssItems for Clients without Customer-Wise Maker Rates
-- Date: 02-Nov-2026
-- Description: Merges split duplicate VendAssItems records caused by auto-assignment
--              when customer-wise rates are not enabled, preserves active rates,
--              repoints revision history, and resets CustCode to 'Stock'.
-- ==============================================================================

IF EXISTS (
    SELECT 1 
    FROM VendAssItems 
    GROUP BY VendID, ProcessID, ItemID 
    HAVING COUNT(*) > 1
)
BEGIN
    BEGIN TRANSACTION;

    -- 1. Identify keepers and redundant duplicates
    -- Keeper is prioritized by: active rate (> 0) first, then latest EntryID
    WITH RankedVAI AS (
        SELECT EntryID, VendID, ProcessID, ItemID, Rate,
               FIRST_VALUE(EntryID) OVER (
                   PARTITION BY VendID, ProcessID, ItemID 
                   ORDER BY CASE WHEN ISNULL(Rate, 0) > 0 THEN 1 ELSE 2 END, EntryID DESC
               ) AS KeeperEntryID
        FROM VendAssItems
    )
    SELECT EntryID AS RedundantID, KeeperEntryID 
    INTO #Duplicates
    FROM RankedVAI
    WHERE EntryID <> KeeperEntryID;

    -- 2. Repoint revision logs from redundant entries to keeper
    UPDATE r
    SET r.VAI_RefID = d.KeeperEntryID
    FROM VendAssItems_Revisions r
    INNER JOIN #Duplicates d ON r.VAI_RefID = d.RedundantID;

    -- 3. Repoint VendAssItems_Rates foreign references if present
    IF OBJECT_ID('VendAssItems_Rates') IS NOT NULL
    BEGIN
        UPDATE vr
        SET vr.VAI_RefID = d.KeeperEntryID
        FROM VendAssItems_Rates vr
        INNER JOIN #Duplicates d ON vr.VAI_RefID = d.RedundantID;
    END

    -- 4. Delete redundant rows
    DELETE vai
    FROM VendAssItems vai
    INNER JOIN #Duplicates d ON vai.EntryID = d.RedundantID;

    -- 5. Standardize CustCode to 'Stock' for clients not using customer-wise rates
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('VendAssItems') AND name = 'CustCode')
    BEGIN
        UPDATE VendAssItems
        SET CustCode = 'Stock'
        WHERE CustCode IS NULL OR CustCode = '' OR CustCode <> 'Stock';
    END

    DROP TABLE #Duplicates;

    COMMIT TRANSACTION;
    PRINT 'VendAssItems duplicates consolidated successfully.';
END
ELSE
BEGIN
    -- Ensure CustCode is standardized to 'Stock' if column exists
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('VendAssItems') AND name = 'CustCode')
    BEGIN
        UPDATE VendAssItems
        SET CustCode = 'Stock'
        WHERE CustCode IS NULL OR CustCode = '' OR CustCode <> 'Stock';
    END
    PRINT 'No duplicate VendAssItems records found. CustCode standardized to Stock.';
END
GO

