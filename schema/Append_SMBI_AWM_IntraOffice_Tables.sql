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

-- 7. CREATE TASK_ASSIGNEES CHILD TABLE
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Task_Assignees')
BEGIN
    CREATE TABLE [dbo].[Task_Assignees] (
        [TaskID] INT NOT NULL,
        [UserID] INT NOT NULL,
        [UserName] NVARCHAR(100) NOT NULL,
        [AssignedAt] DATETIME2 NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_Task_Assignees] PRIMARY KEY CLUSTERED ([TaskID], [UserID]),
        CONSTRAINT [FK_Task_Assignees_TaskItems] FOREIGN KEY ([TaskID]) REFERENCES [TaskItems]([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_Task_Assignees_Users] FOREIGN KEY ([UserID]) REFERENCES [Users]([UserID])
    );
    CREATE NONCLUSTERED INDEX [IX_Task_Assignees_User] ON [Task_Assignees] ([UserID], [UserName]);
    PRINT 'Created table: Task_Assignees';
END
GO

-- 8. TASKITEMS WORKFLOW METADATA COLUMNS
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'SourceEntityType')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [SourceEntityType] NVARCHAR(50) NULL;
    ALTER TABLE [dbo].[TaskItems] ADD [SourceEntityRefId] NVARCHAR(100) NULL;
    ALTER TABLE [dbo].[TaskItems] ADD [TargetRole] NVARCHAR(50) NULL;
    ALTER TABLE [dbo].[TaskItems] ADD [CompletedBy] NVARCHAR(100) NULL;
    ALTER TABLE [dbo].[TaskItems] ADD [ActionUrl] NVARCHAR(300) NULL;
    CREATE NONCLUSTERED INDEX [IX_TaskItems_SourceEntity] ON [dbo].[TaskItems]([SourceEntityType], [SourceEntityRefId]);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_TargetRole_Status] ON [dbo].[TaskItems]([TargetRole], [Status]);
    PRINT 'Added workflow columns to TaskItems';
END
GO

-- 9. CREATE TASK_ROLES CHILD TABLE (FOR MULTI-ROLE ASSIGNMENT)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Task_Roles')
BEGIN
    CREATE TABLE [dbo].[Task_Roles] (
        [TaskID] INT NOT NULL,
        [RoleName] NVARCHAR(50) NOT NULL,
        [AssignedAt] DATETIME2 NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_Task_Roles] PRIMARY KEY CLUSTERED ([TaskID], [RoleName]),
        CONSTRAINT [FK_Task_Roles_TaskItems] FOREIGN KEY ([TaskID]) REFERENCES [TaskItems]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_Task_Roles_Role] ON [Task_Roles] ([RoleName]);
    PRINT 'Created table: Task_Roles';
END
GO

-- 10. CREATE APPNOTIFICATIONS TABLE (FOR PERSISTENT NOTIFICATIONS & ALERTS)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AppNotifications')
BEGIN
    CREATE TABLE [dbo].[AppNotifications] (
        [Id] NVARCHAR(50) NOT NULL PRIMARY KEY,
        [Category] INT NOT NULL DEFAULT 5, -- 0:Message, 1:Announcement, 2:Minute, 3:Meeting, 4:Task, 5:System
        [Title] NVARCHAR(250) NOT NULL,
        [Message] NVARCHAR(MAX) NOT NULL,
        [SenderName] NVARCHAR(100) NULL,
        [TargetUserId] NVARCHAR(100) NULL,
        [ActionUrl] NVARCHAR(500) NOT NULL DEFAULT '/',
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [IsRead] BIT NOT NULL DEFAULT 0,
        [IsReadReceipt] BIT NOT NULL DEFAULT 0
    );
    CREATE NONCLUSTERED INDEX [IX_AppNotifications_TargetUser_IsRead] 
        ON [dbo].[AppNotifications] ([TargetUserId], [IsRead], [CreatedAt] DESC);
    PRINT 'Created table: AppNotifications';
END
GO

-- 11. CREATE PROCESSGROUP_HUB_SUPERVISORS CHILD TABLE
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProcessGroup_Hub_Supervisors')
BEGIN
    CREATE TABLE [dbo].[ProcessGroup_Hub_Supervisors] (
        [GroupID] INT NOT NULL,
        [Hub_Name] VARCHAR(50) NOT NULL,
        [UserID] INT NOT NULL,
        [UserName] NVARCHAR(100) NOT NULL,
        [AssignedAt] DATETIME2 NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_ProcessGroup_Hub_Supervisors] PRIMARY KEY CLUSTERED ([GroupID], [Hub_Name], [UserID]),
        CONSTRAINT [FK_PG_HubSupervisors_ProcessGroups] FOREIGN KEY ([GroupID]) 
            REFERENCES [ProcessGroups]([EntryID]) ON DELETE CASCADE,
        CONSTRAINT [FK_PG_HubSupervisors_HubNames] FOREIGN KEY ([Hub_Name]) 
            REFERENCES [Hub_Names]([Hub_Name]) ON UPDATE CASCADE,
        CONSTRAINT [FK_PG_HubSupervisors_Users] FOREIGN KEY ([UserID]) 
            REFERENCES [Users]([UserID])
    );
    CREATE NONCLUSTERED INDEX [IX_PG_HubSupervisors_Group] 
        ON [ProcessGroup_Hub_Supervisors] ([GroupID], [Hub_Name]);
    CREATE NONCLUSTERED INDEX [IX_PG_HubSupervisors_User] 
        ON [ProcessGroup_Hub_Supervisors] ([UserID], [UserName]);
    PRINT 'Created table: ProcessGroup_Hub_Supervisors';
END
GO

-- 12. CREATE PPC ORDER PLANNING TABLES
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PPC_Order_Planning_Master')
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Planning_Master] (
        [OrderNo] VARCHAR(50) NOT NULL PRIMARY KEY,
        [PlannedBy] NVARCHAR(100) NOT NULL,
        [PlannedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [Status] VARCHAR(20) NOT NULL DEFAULT 'Planned',
        [Notes] NVARCHAR(500) NULL,
        CONSTRAINT [FK_PpcPlanningMaster_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE
    );
    PRINT 'Created table: PPC_Order_Planning_Master';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PPC_Order_Item_Planning')
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Planning] (
        [EntryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [OrderNo] VARCHAR(50) NOT NULL,
        [ItemID] VARCHAR(50) NOT NULL,
        [OrderQty] INT NOT NULL,
        [StockQty] INT NOT NULL DEFAULT 0,
        [TotalPurchaseQty] INT NOT NULL DEFAULT 0,
        [ProductionQty] INT NOT NULL DEFAULT 0,
        [Remarks] NVARCHAR(250) NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [CreatedBy] NVARCHAR(100) NOT NULL,
        CONSTRAINT [FK_PpcItemPlanning_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_PpcItemPlanning_Order_Item] 
        ON [dbo].[PPC_Order_Item_Planning] ([OrderNo], [ItemID]);
    PRINT 'Created table: PPC_Order_Item_Planning';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PPC_Order_Item_Purchases')
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Purchases] (
        [EntryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [OrderNo] VARCHAR(50) NOT NULL,
        [ItemID] VARCHAR(50) NOT NULL,
        [ProcessID] INT NOT NULL,
        [VendID] INT NULL,
        [PurchaseQty] INT NOT NULL,
        [PurchaseRate] FLOAT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [CreatedBy] NVARCHAR(100) NOT NULL,
        CONSTRAINT [FK_PpcItemPurchases_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_PpcItemPurchases_Order_Item] 
        ON [dbo].[PPC_Order_Item_Purchases] ([OrderNo], [ItemID]);
    PRINT 'Created table: PPC_Order_Item_Purchases';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PPC_Order_Item_Hub_Schedules')
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Hub_Schedules] (
        [EntryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [OrderNo] VARCHAR(50) NOT NULL,
        [ItemID] VARCHAR(50) NOT NULL,
        [Hub_Name] VARCHAR(50) NOT NULL,
        [PlannedQty] INT NOT NULL DEFAULT 0,
        [StartDate] DATETIME NOT NULL,
        [EndDate] DATETIME NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [CreatedBy] NVARCHAR(100) NOT NULL,
        CONSTRAINT [FK_PpcHubSchedules_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_PpcHubSchedules_Order_Item] 
        ON [dbo].[PPC_Order_Item_Hub_Schedules] ([OrderNo], [ItemID]);
    PRINT 'Created table: PPC_Order_Item_Hub_Schedules';
END
GO

PRINT '==============================================================================================';
PRINT 'SMBI_AWM IntraOffice & PPC Planning schema appended successfully!';
PRINT '==============================================================================================';
GO

