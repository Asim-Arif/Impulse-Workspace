-- Table: dbo.AttendanceMachine
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.AttendanceMachine') AND name = 'Clientkey')
    ALTER TABLE dbo.AttendanceMachine ADD [Clientkey] varchar(50) NULL DEFAULT ('Default');
GO