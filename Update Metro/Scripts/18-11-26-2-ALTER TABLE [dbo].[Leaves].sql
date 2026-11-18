-- =====================================================================
-- Add HalfDayType column to dbo.Leaves (1 = First Half, 2 = Second Half)
-- =====================================================================
IF OBJECT_ID('dbo.Leaves', 'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.Leaves', 'HalfDayType') IS NULL
    BEGIN
        ALTER TABLE [dbo].[Leaves] ADD [HalfDayType] TINYINT NULL;
    END
END
GO