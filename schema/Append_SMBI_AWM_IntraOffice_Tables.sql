-- ==============================================================================================
-- IntraOffice & CRM Suite - Append Script for SMBI_AWM
-- Target: Microsoft SQL Server
-- Appends Leads, LeadActivities, DownloadItems, EmailConfigurations, TaskStatusHistories,
-- and missing columns on existing collaboration tables.
-- Direct integration with [Users] (UserName) and [ForeignCustomers] (CustCode).
-- ==============================================================================================

USE [SMBI_AWM];
GO

-- 1. UPDATE EXISTING TABLES WITH NEW COLUMNS

-- Meetings: Voice Note Path
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('Meetings') AND name = 'VoiceNotePath')
BEGIN
    ALTER TABLE [dbo].[Meetings] ADD [VoiceNotePath] NVARCHAR(500) NULL;
    PRINT 'Added column VoiceNotePath to Meetings';
END
GO

-- MinuteApprovals: Read Receipts
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('MinuteApprovals') AND name = 'IsRead')
BEGIN
    ALTER TABLE [dbo].[MinuteApprovals] ADD [IsRead] BIT NOT NULL CONSTRAINT [DF_MinuteApprovals_IsRead] DEFAULT 0;
    PRINT 'Added column IsRead to MinuteApprovals';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('MinuteApprovals') AND name = 'ReadAt')
BEGIN
    ALTER TABLE [dbo].[MinuteApprovals] ADD [ReadAt] DATETIME2 NULL;
    PRINT 'Added column ReadAt to MinuteApprovals';
END
GO

-- TaskItems: Multi-assignees, Read state, Started timestamp, Email sent flag
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'AdditionalAssigneeIds')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [AdditionalAssigneeIds] NVARCHAR(MAX) NULL;
    PRINT 'Added column AdditionalAssigneeIds to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'AssignedToNames')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [AssignedToNames] NVARCHAR(MAX) NULL;
    PRINT 'Added column AssignedToNames to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'EmailMessageSent')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [EmailMessageSent] BIT NOT NULL CONSTRAINT [DF_TaskItems_EmailSent] DEFAULT 0;
    PRINT 'Added column EmailMessageSent to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'IsRead')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [IsRead] BIT NOT NULL CONSTRAINT [DF_TaskItems_IsRead] DEFAULT 0;
    PRINT 'Added column IsRead to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'StartedAt')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [StartedAt] DATETIME2 NULL;
    PRINT 'Added column StartedAt to TaskItems';
END
GO


-- 2. CREATE LEADS TABLE (Integrated with ForeignCustomers)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Leads')
BEGIN
    CREATE TABLE [dbo].[Leads] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [LeadNumber] NVARCHAR(50) NOT NULL,
        [CompanyName] NVARCHAR(200) NOT NULL,
        [ContactPerson] NVARCHAR(150) NOT NULL,
        [Designation] NVARCHAR(100) NULL,
        [Email] NVARCHAR(150) NULL,
        [Phone] NVARCHAR(50) NULL,
        [WhatsApp] NVARCHAR(50) NULL,
        [Country] NVARCHAR(100) NOT NULL,
        [City] NVARCHAR(100) NULL,
        [Website] NVARCHAR(200) NULL,
        [Source] NVARCHAR(100) NOT NULL,
        [Industry] NVARCHAR(100) NULL,
        [ProductInterest] NVARCHAR(200) NULL,
        [EstimatedQuantity] INT NOT NULL DEFAULT 0,
        [EstimatedValue] DECIMAL(18,2) NOT NULL DEFAULT 0,
        [Currency] NVARCHAR(10) NOT NULL DEFAULT 'USD',
        [ExpectedOrderDate] DATETIME2 NULL,
        [Status] NVARCHAR(50) NOT NULL DEFAULT 'New', -- New, Contacted, Qualified, Proposal, Won, Lost, Converted
        [Priority] NVARCHAR(20) NOT NULL DEFAULT 'Medium', -- Low, Medium, High, Urgent
        [Notes] NVARCHAR(MAX) NULL,
        [NextFollowUpDate] DATETIME2 NULL,
        [ConvertedCustCode] VARCHAR(50) NULL, -- References ForeignCustomers.CustCode
        [AssignedTo] NVARCHAR(100) NULL,      -- References Users.UserName
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [UpdatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_Leads_Status] ON [dbo].[Leads]([Status]);
    CREATE NONCLUSTERED INDEX [IX_Leads_ConvertedCustCode] ON [dbo].[Leads]([ConvertedCustCode]);
    CREATE NONCLUSTERED INDEX [IX_Leads_AssignedTo] ON [dbo].[Leads]([AssignedTo]);
    PRINT 'Created table: Leads';
END
GO

-- 3. CREATE LEAD ACTIVITIES TABLE
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'LeadActivities')
BEGIN
    CREATE TABLE [dbo].[LeadActivities] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [LeadId] INT NOT NULL,
        [ActivityType] NVARCHAR(50) NOT NULL, -- Call, Email, Meeting, WhatsApp, Note, Site Visit
        [Description] NVARCHAR(MAX) NOT NULL,
        [PerformedBy] NVARCHAR(100) NOT NULL,  -- References Users.UserName
        [ActivityDate] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [NextFollowUpDate] DATETIME2 NULL,
        CONSTRAINT [FK_LeadActivities_Leads] FOREIGN KEY ([LeadId]) REFERENCES [dbo].[Leads]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_LeadActivities_LeadId] ON [dbo].[LeadActivities]([LeadId]);
    PRINT 'Created table: LeadActivities';
END
GO

-- 4. CREATE TASK STATUS HISTORIES TABLE
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TaskStatusHistories')
BEGIN
    CREATE TABLE [dbo].[TaskStatusHistories] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [TaskId] INT NOT NULL,
        [OldStatus] INT NOT NULL,
        [NewStatus] INT NOT NULL,
        [ChangedBy] NVARCHAR(100) NOT NULL,
        [ChangedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [Notes] NVARCHAR(MAX) NULL,
        CONSTRAINT [FK_TaskStatusHistories_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_TaskStatusHistories_TaskId] ON [dbo].[TaskStatusHistories]([TaskId]);
    PRINT 'Created table: TaskStatusHistories';
END
GO

-- 5. CREATE DOWNLOAD ITEMS TABLE
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'DownloadItems')
BEGIN
    CREATE TABLE [dbo].[DownloadItems] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Title] NVARCHAR(200) NOT NULL,
        [Description] NVARCHAR(1000) NULL,
        [Category] NVARCHAR(100) NOT NULL DEFAULT 'General',
        [FileName] NVARCHAR(500) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT 0,
        [ContentType] NVARCHAR(100) NULL,
        [Version] NVARCHAR(50) NOT NULL DEFAULT '1.0',
        [DownloadCount] INT NOT NULL DEFAULT 0,
        [DepartmentId] VARCHAR(50) NULL,
        [CreatedBy] NVARCHAR(100) NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    PRINT 'Created table: DownloadItems';
END
GO

-- 6. CREATE EMAIL CONFIGURATIONS TABLE
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'EmailConfigurations')
BEGIN
    CREATE TABLE [dbo].[EmailConfigurations] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [SmtpServer] NVARCHAR(200) NOT NULL,
        [SmtpPort] INT NOT NULL DEFAULT 587,
        [SenderEmail] NVARCHAR(200) NOT NULL,
        [SenderName] NVARCHAR(200) NOT NULL,
        [Username] NVARCHAR(200) NOT NULL,
        [EncryptedPassword] NVARCHAR(500) NOT NULL,
        [EnableSsl] BIT NOT NULL DEFAULT 1,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    PRINT 'Created table: EmailConfigurations';
END
GO

PRINT '==============================================================================================';
PRINT 'SMBI_AWM IntraOffice & Leads schema appended successfully!';
PRINT '==============================================================================================';
GO
