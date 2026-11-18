-- =====================================================================
-- Add AttendanceMachineReadLog table for IFace machine read auditing
-- =====================================================================
IF OBJECT_ID('dbo.AttendanceMachineReadLog', 'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[AttendanceMachineReadLog] (
        [EntryID]          INT IDENTITY(1,1) NOT NULL CONSTRAINT [PK_AttendanceMachineReadLog] PRIMARY KEY,
        [UserName]         VARCHAR(100) NULL,
        [ClientIPAddress]  VARCHAR(50)  NULL,
        [MachineIPAddress] VARCHAR(50)  NULL,
        [MachineNo]        INT          NULL,
        [ReadStartTime]    DATETIME     NOT NULL,
        [ReadEndTime]      DATETIME     NOT NULL,
        [TotalRecordsRead] INT          NOT NULL CONSTRAINT [DF_AttMachineReadLog_Read] DEFAULT (0),
        [TotalTransferred] INT          NOT NULL CONSTRAINT [DF_AttMachineReadLog_Trans] DEFAULT (0),
        [Remarks]          VARCHAR(500) NULL
    );
END
GO
