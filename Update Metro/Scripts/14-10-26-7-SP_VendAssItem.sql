CREATE OR ALTER PROCEDURE [dbo].[SP_VendAssItem] 
    @VendID AS INT,
    @ItemID AS VARCHAR(50) = '0',
    @ProcessID AS INT,
    @Rate AS REAL = 0,
    @CustCode AS VARCHAR(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @PlantRate AS FLOAT;
    DECLARE @StampRate AS FLOAT;
    DECLARE @SnaffRate AS FLOAT;
    DECLARE @Unit AS VARCHAR(50);
    DECLARE @EffectiveCustCode AS VARCHAR(50) = ISNULL(NULLIF(@CustCode, ''), 'Stock');

    IF @ItemID <> '0' 
    BEGIN
        SELECT @PlantRate = PlantRate, @StampRate = StampRate, @SnaffRate = SnaffRate, @Unit = Unit 
        FROM Items 
        WHERE ItemID = @ItemID;

        INSERT INTO VendAssItems(VendID, ProcessID, ItemID, Rate, PlantRate, StampRate, SnaffRate, Unit, CustCode)
        VALUES(@VendID, @ProcessID, @ItemID, @Rate, @PlantRate, @StampRate, @SnaffRate, @Unit, @EffectiveCustCode);
    END
    ELSE
    BEGIN
        -- Bulk assign unassigned items for this maker, process & customer
        INSERT INTO VendAssItems(VendID, ProcessID, ItemID, PlantRate, StampRate, SnaffRate, CustCode)
        SELECT @VendID, @ProcessID, ItemID, PlantRate, StampRate, SnaffRate, @EffectiveCustCode
        FROM Items 
        WHERE ItemID IN (SELECT ItemID FROM ItemProcesses WHERE ProcessID = @ProcessID) 
          AND ItemID NOT IN (
              SELECT ItemID 
              FROM VendAssItems 
              WHERE VendID = @VendID 
                AND ProcessID = @ProcessID
                AND (CustCode = @EffectiveCustCode OR (@EffectiveCustCode = 'Stock' AND (CustCode IS NULL OR CustCode = '' OR CustCode = 'Stock')))
          );
    END
END
GO