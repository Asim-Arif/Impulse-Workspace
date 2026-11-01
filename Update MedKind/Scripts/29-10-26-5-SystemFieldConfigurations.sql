
-- ==============================================================================
-- SystemFieldConfigurations Table & Initial Seed
-- Universal configurable form field rules (e.g. Mandatory / Optional / Visible)
-- ==============================================================================
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'SystemFieldConfigurations')
BEGIN
    CREATE TABLE [dbo].[SystemFieldConfigurations] (
        [SettingID] INT IDENTITY(1,1) PRIMARY KEY,
        [Module] VARCHAR(50) NOT NULL,
        [FormKey] VARCHAR(50) NOT NULL,
        [FieldKey] VARCHAR(50) NOT NULL,
        [IsRequired] BIT NOT NULL DEFAULT 1,
        [IsVisible] BIT NOT NULL DEFAULT 1,
        [Description] VARCHAR(255) NULL,
        CONSTRAINT UQ_SystemFieldConfigurations UNIQUE ([Module], [FormKey], [FieldKey])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[SystemFieldConfigurations] WHERE [Module] = 'Production' AND [FormKey] = 'LotIssuance' AND [FieldKey] = 'FactoryWorker')
BEGIN
    INSERT INTO [dbo].[SystemFieldConfigurations] ([Module], [FormKey], [FieldKey], [IsRequired], [IsVisible], [Description])
    VALUES ('Production', 'LotIssuance', 'FactoryWorker', 1, 1, 'Factory Worker selection on Lot Issuance');
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[SystemFieldConfigurations] WHERE [Module] = 'Production' AND [FormKey] = 'LotIssuance' AND [FieldKey] = 'CountedBy')
BEGIN
    INSERT INTO [dbo].[SystemFieldConfigurations] ([Module], [FormKey], [FieldKey], [IsRequired], [IsVisible], [Description])
    VALUES ('Production', 'LotIssuance', 'CountedBy', 0, 1, 'Counted By selection on Lot Issuance');
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[SystemFieldConfigurations] WHERE [Module] = 'Production' AND [FormKey] = 'LotReceiving' AND [FieldKey] = 'FactoryWorker')
BEGIN
    INSERT INTO [dbo].[SystemFieldConfigurations] ([Module], [FormKey], [FieldKey], [IsRequired], [IsVisible], [Description])
    VALUES ('Production', 'LotReceiving', 'FactoryWorker', 1, 1, 'Factory Worker selection on Lot Receiving');
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[SystemFieldConfigurations] WHERE [Module] = 'Production' AND [FormKey] = 'LotReceiving' AND [FieldKey] = 'CheckedBy')
BEGIN
    INSERT INTO [dbo].[SystemFieldConfigurations] ([Module], [FormKey], [FieldKey], [IsRequired], [IsVisible], [Description])
    VALUES ('Production', 'LotReceiving', 'CheckedBy', 1, 1, 'Checked By selection on Lot Receiving');
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[SystemFieldConfigurations] WHERE [Module] = 'Production' AND [FormKey] = 'LotReceiving' AND [FieldKey] = 'CountedBy')
BEGIN
    INSERT INTO [dbo].[SystemFieldConfigurations] ([Module], [FormKey], [FieldKey], [IsRequired], [IsVisible], [Description])
    VALUES ('Production', 'LotReceiving', 'CountedBy', 1, 1, 'Line Counted By selection on Lot Receiving');
END
GO
