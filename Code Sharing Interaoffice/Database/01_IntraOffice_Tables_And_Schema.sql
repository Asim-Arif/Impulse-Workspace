-- ====================================================================================
-- IntraOffice Collaboration & CRM Module - Standalone Database Schema Migration
-- Designed for: Impulse ERP Ecosystem (Surgical, Textile, and Manufacturing Suites)
-- Execution: 100% Idempotent, Safe for Repeated Runs, Independent of Database Name
-- ====================================================================================

SET NOCOUNT ON;

PRINT 'Starting IntraOffice Module Schema Migration...';

-- ====================================================================================
-- 1. USER PRESENCE & REALTIME STATUS
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'UserPresences')
BEGIN
    CREATE TABLE [dbo].[UserPresences] (
        [UserId] NVARCHAR(100) NOT NULL,
        [Status] INT NOT NULL DEFAULT (0),
        [LastSeen] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        [ConnectionId] NVARCHAR(400) NULL,
        CONSTRAINT [PK_UserPresences] PRIMARY KEY ([UserId] ASC)
    );
    PRINT '  -> Table [UserPresences] created.';
END
GO

-- ====================================================================================
-- 2. CHANNELS & WORKSPACE TEAMS
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Channels')
BEGIN
    CREATE TABLE [dbo].[Channels] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [Name] NVARCHAR(200) NOT NULL,
        [Description] NVARCHAR(1000) NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [IsPrivate] BIT NOT NULL DEFAULT (0),
        [IsActive] BIT NOT NULL DEFAULT (1),
        [CreatedBy] NVARCHAR(100) NOT NULL,
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_Channels] PRIMARY KEY ([Id] ASC)
    );
    CREATE NONCLUSTERED INDEX [IX_Channels_DepartmentId] ON [dbo].[Channels] ([DepartmentId] ASC);
    CREATE NONCLUSTERED INDEX [IX_Channels_CreatedBy] ON [dbo].[Channels] ([CreatedBy] ASC);
    PRINT '  -> Table [Channels] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ChannelMembers')
BEGIN
    CREATE TABLE [dbo].[ChannelMembers] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [ChannelId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [IsAdmin] BIT NOT NULL DEFAULT (0),
        [JoinedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_ChannelMembers] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_ChannelMembers_Channels] FOREIGN KEY ([ChannelId]) REFERENCES [dbo].[Channels] ([Id]) ON DELETE CASCADE
    );
    CREATE UNIQUE NONCLUSTERED INDEX [IX_ChannelMembers_ChannelId_UserId] ON [dbo].[ChannelMembers] ([ChannelId] ASC, [UserId] ASC);
    CREATE NONCLUSTERED INDEX [IX_ChannelMembers_UserId] ON [dbo].[ChannelMembers] ([UserId] ASC);
    PRINT '  -> Table [ChannelMembers] created.';
END
GO

-- ====================================================================================
-- 3. MESSAGES & DIRECT CHAT
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Messages')
BEGIN
    CREATE TABLE [dbo].[Messages] (
        [Id] BIGINT IDENTITY(1,1) NOT NULL,
        [ChannelId] INT NULL,
        [SenderId] NVARCHAR(100) NOT NULL,
        [ReceiverId] NVARCHAR(100) NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [MessageType] INT NOT NULL DEFAULT (0),
        [IsRead] BIT NOT NULL DEFAULT (0),
        [IsDeleted] BIT NOT NULL DEFAULT (0),
        [ParentMessageId] BIGINT NULL,
        [SentAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        [EditedAt] DATETIME2(7) NULL,
        CONSTRAINT [PK_Messages] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_Messages_Channels] FOREIGN KEY ([ChannelId]) REFERENCES [dbo].[Channels] ([Id]) ON DELETE SET NULL,
        CONSTRAINT [FK_Messages_ParentMessage] FOREIGN KEY ([ParentMessageId]) REFERENCES [dbo].[Messages] ([Id])
    );
    CREATE NONCLUSTERED INDEX [IX_Messages_ChannelId_SentAt] ON [dbo].[Messages] ([ChannelId] ASC, [SentAt] DESC);
    CREATE NONCLUSTERED INDEX [IX_Messages_Sender_Receiver] ON [dbo].[Messages] ([SenderId] ASC, [ReceiverId] ASC, [SentAt] DESC);
    CREATE NONCLUSTERED INDEX [IX_Messages_ParentMessageId] ON [dbo].[Messages] ([ParentMessageId] ASC);
    PRINT '  -> Table [Messages] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MessageAttachments')
BEGIN
    CREATE TABLE [dbo].[MessageAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [MessageId] BIGINT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT (0),
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_MessageAttachments] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_MessageAttachments_Messages] FOREIGN KEY ([MessageId]) REFERENCES [dbo].[Messages] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MessageAttachments_MessageId] ON [dbo].[MessageAttachments] ([MessageId] ASC);
    PRINT '  -> Table [MessageAttachments] created.';
END
GO

-- ====================================================================================
-- 4. ANNOUNCEMENTS & COMPANY NOTICE BOARD
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Announcements')
BEGIN
    CREATE TABLE [dbo].[Announcements] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [Title] NVARCHAR(400) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [Priority] INT NOT NULL DEFAULT (0),
        [CreatedBy] NVARCHAR(100) NOT NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [IsPinned] BIT NOT NULL DEFAULT (0),
        [IsActive] BIT NOT NULL DEFAULT (1),
        [IsAcknowledged] BIT NOT NULL DEFAULT (0),
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        [UpdatedAt] DATETIME2(7) NULL,
        [ExpiresAt] DATETIME2(7) NULL,
        CONSTRAINT [PK_Announcements] PRIMARY KEY ([Id] ASC)
    );
    CREATE NONCLUSTERED INDEX [IX_Announcements_CreatedBy] ON [dbo].[Announcements] ([CreatedBy] ASC);
    CREATE NONCLUSTERED INDEX [IX_Announcements_DepartmentId] ON [dbo].[Announcements] ([DepartmentId] ASC);
    CREATE NONCLUSTERED INDEX [IX_Announcements_CreatedAt] ON [dbo].[Announcements] ([CreatedAt] DESC);
    PRINT '  -> Table [Announcements] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'AnnouncementAcknowledgments')
BEGIN
    CREATE TABLE [dbo].[AnnouncementAcknowledgments] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [AnnouncementId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [AcknowledgedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_AnnouncementAcknowledgments] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_AnnouncementAcknowledgments_Announcements] FOREIGN KEY ([AnnouncementId]) REFERENCES [dbo].[Announcements] ([Id]) ON DELETE CASCADE
    );
    CREATE UNIQUE NONCLUSTERED INDEX [IX_AnnouncementAcknowledgments_Announcement_User] ON [dbo].[AnnouncementAcknowledgments] ([AnnouncementId] ASC, [UserId] ASC);
    PRINT '  -> Table [AnnouncementAcknowledgments] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'AnnouncementAttachments')
BEGIN
    CREATE TABLE [dbo].[AnnouncementAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [AnnouncementId] INT NOT NULL,
        [FileName] NVARCHAR(MAX) NOT NULL,
        [FilePath] NVARCHAR(MAX) NOT NULL,
        [FileSize] BIGINT NOT NULL,
        [ContentType] NVARCHAR(MAX) NULL,
        [UploadedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_AnnouncementAttachments] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_AnnouncementAttachments_Announcements] FOREIGN KEY ([AnnouncementId]) REFERENCES [dbo].[Announcements] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_AnnouncementAttachments_AnnouncementId] ON [dbo].[AnnouncementAttachments] ([AnnouncementId] ASC);
    PRINT '  -> Table [AnnouncementAttachments] created.';
END
GO

-- ====================================================================================
-- 5. TASK MANAGEMENT & WORKFLOW ORCHESTRATION
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'TaskItems')
BEGIN
    CREATE TABLE [dbo].[TaskItems] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [Title] NVARCHAR(400) NOT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [AssignedTo] NVARCHAR(100) NULL,
        [AssignedBy] NVARCHAR(100) NOT NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [Priority] INT NOT NULL DEFAULT (1),
        [Status] INT NOT NULL DEFAULT (0),
        [DueDate] DATETIME2(7) NULL,
        [WhatsAppMessageSent] BIT NOT NULL DEFAULT (0),
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        [UpdatedAt] DATETIME2(7) NULL,
        [CompletedAt] DATETIME2(7) NULL,
        [AdditionalAssigneeIds] NVARCHAR(MAX) NULL,
        [AssignedToNames] NVARCHAR(MAX) NULL,
        [EmailMessageSent] BIT NOT NULL DEFAULT (0),
        [IsRead] BIT NOT NULL DEFAULT (0),
        [StartedAt] DATETIME2(7) NULL,
        [SourceEntityType] NVARCHAR(50) NULL,
        [SourceEntityRefId] NVARCHAR(100) NULL,
        [TargetRole] NVARCHAR(50) NULL,
        [CompletedBy] NVARCHAR(100) NULL,
        [ActionUrl] NVARCHAR(300) NULL,
        [DueWarningSent] BIT NOT NULL DEFAULT (0),
        [OverdueWarningSent] BIT NOT NULL DEFAULT (0),
        CONSTRAINT [PK_TaskItems] PRIMARY KEY ([Id] ASC)
    );
    CREATE NONCLUSTERED INDEX [IX_TaskItems_AssignedTo] ON [dbo].[TaskItems] ([AssignedTo] ASC);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_AssignedBy] ON [dbo].[TaskItems] ([AssignedBy] ASC);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_Status] ON [dbo].[TaskItems] ([Status] ASC);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_DepartmentId] ON [dbo].[TaskItems] ([DepartmentId] ASC);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_SourceEntity] ON [dbo].[TaskItems] ([SourceEntityType] ASC, [SourceEntityRefId] ASC);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_TargetRole_Status] ON [dbo].[TaskItems] ([TargetRole] ASC, [Status] ASC);
    PRINT '  -> Table [TaskItems] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Task_Assignees')
BEGIN
    CREATE TABLE [dbo].[Task_Assignees] (
        [TaskID] INT NOT NULL,
        [UserID] INT NOT NULL,
        [UserName] NVARCHAR(100) NOT NULL,
        [AssignedAt] DATETIME2(7) NOT NULL DEFAULT (GETDATE()),
        CONSTRAINT [PK_Task_Assignees] PRIMARY KEY ([TaskID] ASC, [UserID] ASC),
        CONSTRAINT [FK_Task_Assignees_TaskItems] FOREIGN KEY ([TaskID]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_Task_Assignees_User] ON [dbo].[Task_Assignees] ([UserID] ASC, [UserName] ASC);
    PRINT '  -> Table [Task_Assignees] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Task_Roles')
BEGIN
    CREATE TABLE [dbo].[Task_Roles] (
        [TaskID] INT NOT NULL,
        [RoleName] NVARCHAR(50) NOT NULL,
        [AssignedAt] DATETIME2(7) NOT NULL DEFAULT (GETDATE()),
        CONSTRAINT [PK_Task_Roles] PRIMARY KEY ([TaskID] ASC, [RoleName] ASC),
        CONSTRAINT [FK_Task_Roles_TaskItems] FOREIGN KEY ([TaskID]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_Task_Roles_Role] ON [dbo].[Task_Roles] ([RoleName] ASC);
    PRINT '  -> Table [Task_Roles] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'TaskComments')
BEGIN
    CREATE TABLE [dbo].[TaskComments] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [TaskId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_TaskComments] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_TaskComments_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_TaskComments_TaskId] ON [dbo].[TaskComments] ([TaskId] ASC);
    PRINT '  -> Table [TaskComments] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'TaskAttachments')
BEGIN
    CREATE TABLE [dbo].[TaskAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [TaskId] INT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT (0),
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_TaskAttachments] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_TaskAttachments_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_TaskAttachments_TaskId] ON [dbo].[TaskAttachments] ([TaskId] ASC);
    PRINT '  -> Table [TaskAttachments] created.';
END
GO

-- ====================================================================================
-- 6. MEETINGS & COLLABORATION SESSIONS
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Meetings')
BEGIN
    CREATE TABLE [dbo].[Meetings] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [Title] NVARCHAR(400) NOT NULL,
        [OrganizerId] NVARCHAR(100) NOT NULL,
        [MeetingUrl] NVARCHAR(1000) NULL,
        [ScheduledStartTime] DATETIME2(7) NOT NULL,
        [ScheduledEndTime] DATETIME2(7) NULL,
        [Status] INT NOT NULL DEFAULT (0),
        [MeetingMinutes] NVARCHAR(MAX) NULL,
        [IsReminderSent] BIT NOT NULL DEFAULT (0),
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        [VoiceNotePath] NVARCHAR(500) NULL,
        CONSTRAINT [PK_Meetings] PRIMARY KEY ([Id] ASC)
    );
    CREATE NONCLUSTERED INDEX [IX_Meetings_OrganizerId] ON [dbo].[Meetings] ([OrganizerId] ASC);
    CREATE NONCLUSTERED INDEX [IX_Meetings_ScheduledStartTime] ON [dbo].[Meetings] ([ScheduledStartTime] ASC);
    PRINT '  -> Table [Meetings] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MeetingParticipants')
BEGIN
    CREATE TABLE [dbo].[MeetingParticipants] (
        [MeetingId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [HasAttended] BIT NOT NULL DEFAULT (0),
        CONSTRAINT [PK_MeetingParticipants] PRIMARY KEY ([MeetingId] ASC, [UserId] ASC),
        CONSTRAINT [FK_MeetingParticipants_Meetings] FOREIGN KEY ([MeetingId]) REFERENCES [dbo].[Meetings] ([Id]) ON DELETE CASCADE
    );
    PRINT '  -> Table [MeetingParticipants] created.';
END
GO

-- ====================================================================================
-- 7. MINUTES APPROVAL WORKFLOW ENGINE
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MinuteTypes')
BEGIN
    CREATE TABLE [dbo].[MinuteTypes] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [Name] NVARCHAR(200) NOT NULL,
        [IsActive] BIT NOT NULL DEFAULT (1),
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_MinuteTypes] PRIMARY KEY ([Id] ASC)
    );
    PRINT '  -> Table [MinuteTypes] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MinuteApprovals')
BEGIN
    CREATE TABLE [dbo].[MinuteApprovals] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [Date] DATETIME2(7) NOT NULL,
        [No] NVARCHAR(100) NOT NULL,
        [Type] NVARCHAR(100) NOT NULL,
        [Subject] NVARCHAR(MAX) NOT NULL,
        [Points] NVARCHAR(MAX) NOT NULL,
        [ForwardToUserId] NVARCHAR(100) NULL,
        [Currency] NVARCHAR(20) NULL,
        [TotalAmount] DECIMAL(18,2) NULL,
        [AdvancePercentage] DECIMAL(5,2) NULL,
        [AdvanceAmount] DECIMAL(18,2) NULL,
        [IsUrgent] BIT NOT NULL DEFAULT (0),
        [CloseByInitiator] BIT NOT NULL DEFAULT (0),
        [CreatedByUserId] NVARCHAR(100) NOT NULL,
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        [Status] NVARCHAR(100) NOT NULL DEFAULT ('Pending'),
        [SignaturePath] NVARCHAR(MAX) NULL,
        [PurchaseOrderValue] DECIMAL(18,2) NULL,
        [PurchaseAdvanceRecommend] DECIMAL(18,2) NULL,
        [PurchaseApprovedAmount] DECIMAL(18,2) NULL,
        [RequestedStockQty] DECIMAL(18,2) NULL,
        [CurrentStockQty] DECIMAL(18,2) NULL,
        [ApprovedStockQty] DECIMAL(18,2) NULL,
        [HRLeaveType] NVARCHAR(510) NULL,
        [FinancialType] NVARCHAR(510) NULL,
        [IsRead] BIT NOT NULL DEFAULT (0),
        [ReadAt] DATETIME2(7) NULL,
        CONSTRAINT [PK_MinuteApprovals] PRIMARY KEY ([Id] ASC)
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_CreatedByUserId] ON [dbo].[MinuteApprovals] ([CreatedByUserId] ASC);
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_ForwardToUserId] ON [dbo].[MinuteApprovals] ([ForwardToUserId] ASC);
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_Date] ON [dbo].[MinuteApprovals] ([Date] DESC);
    PRINT '  -> Table [MinuteApprovals] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MinuteWorkflowHistories')
BEGIN
    CREATE TABLE [dbo].[MinuteWorkflowHistories] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [MinuteApprovalId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [ActionTaken] NVARCHAR(100) NOT NULL,
        [Remarks] NVARCHAR(MAX) NULL,
        [SignaturePath] NVARCHAR(MAX) NULL,
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_MinuteWorkflowHistories] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_MinuteWorkflowHistories_MinuteApprovals] FOREIGN KEY ([MinuteApprovalId]) REFERENCES [dbo].[MinuteApprovals] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteWorkflowHistories_MinuteApprovalId] ON [dbo].[MinuteWorkflowHistories] ([MinuteApprovalId] ASC);
    PRINT '  -> Table [MinuteWorkflowHistories] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MinuteAttachments')
BEGIN
    CREATE TABLE [dbo].[MinuteAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [MinuteApprovalId] INT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT (0),
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_MinuteAttachments] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_MinuteAttachments_MinuteApprovals] FOREIGN KEY ([MinuteApprovalId]) REFERENCES [dbo].[MinuteApprovals] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteAttachments_MinuteApprovalId] ON [dbo].[MinuteAttachments] ([MinuteApprovalId] ASC);
    PRINT '  -> Table [MinuteAttachments] created.';
END
GO

-- ====================================================================================
-- 8. STICKY NOTES
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'StickyNotes')
BEGIN
    CREATE TABLE [dbo].[StickyNotes] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [Color] NVARCHAR(100) NOT NULL DEFAULT ('#fffa65'),
        [XPos] INT NOT NULL DEFAULT (100),
        [YPos] INT NOT NULL DEFAULT (100),
        [ReminderTime] DATETIME2(7) NULL,
        [IsReminderSent] BIT NOT NULL DEFAULT (0),
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_StickyNotes] PRIMARY KEY ([Id] ASC)
    );
    CREATE NONCLUSTERED INDEX [IX_StickyNotes_UserId] ON [dbo].[StickyNotes] ([UserId] ASC);
    PRINT '  -> Table [StickyNotes] created.';
END
GO

-- ====================================================================================
-- 9. LEADS & PRE-SALES CRM PIPELINE
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Leads')
BEGIN
    CREATE TABLE [dbo].[Leads] (
        [Id] INT IDENTITY(1,1) NOT NULL,
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
        [EstimatedQuantity] INT NOT NULL DEFAULT (0),
        [EstimatedValue] DECIMAL(18,2) NOT NULL DEFAULT (0),
        [Currency] NVARCHAR(10) NOT NULL DEFAULT ('USD'),
        [ExpectedOrderDate] DATETIME2(7) NULL,
        [Status] NVARCHAR(50) NOT NULL DEFAULT ('New'),
        [Priority] NVARCHAR(20) NOT NULL DEFAULT ('Medium'),
        [Notes] NVARCHAR(MAX) NULL,
        [NextFollowUpDate] DATETIME2(7) NULL,
        [ConvertedCustCode] VARCHAR(50) NULL,
        [AssignedTo] NVARCHAR(100) NULL,
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        [UpdatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_Leads] PRIMARY KEY ([Id] ASC)
    );
    CREATE NONCLUSTERED INDEX [IX_Leads_Status] ON [dbo].[Leads] ([Status] ASC);
    CREATE NONCLUSTERED INDEX [IX_Leads_ConvertedCustCode] ON [dbo].[Leads] ([ConvertedCustCode] ASC);
    CREATE NONCLUSTERED INDEX [IX_Leads_AssignedTo] ON [dbo].[Leads] ([AssignedTo] ASC);
    PRINT '  -> Table [Leads] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'LeadActivities')
BEGIN
    CREATE TABLE [dbo].[LeadActivities] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [LeadId] INT NOT NULL,
        [ActivityType] NVARCHAR(50) NOT NULL,
        [Description] NVARCHAR(MAX) NOT NULL,
        [PerformedBy] NVARCHAR(100) NOT NULL,
        [ActivityDate] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        [NextFollowUpDate] DATETIME2(7) NULL,
        CONSTRAINT [PK_LeadActivities] PRIMARY KEY ([Id] ASC),
        CONSTRAINT [FK_LeadActivities_Leads] FOREIGN KEY ([LeadId]) REFERENCES [dbo].[Leads] ([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_LeadActivities_LeadId] ON [dbo].[LeadActivities] ([LeadId] ASC);
    PRINT '  -> Table [LeadActivities] created.';
END
GO

-- ====================================================================================
-- 10. EMAIL CONFIGURATION & CRM TEMPLATES
-- ====================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EmailConfigurations')
BEGIN
    CREATE TABLE [dbo].[EmailConfigurations] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [SmtpServer] NVARCHAR(200) NOT NULL,
        [SmtpPort] INT NOT NULL DEFAULT (587),
        [SenderEmail] NVARCHAR(200) NOT NULL,
        [SenderName] NVARCHAR(200) NOT NULL,
        [Username] NVARCHAR(200) NOT NULL,
        [EncryptedPassword] NVARCHAR(500) NOT NULL,
        [EnableSsl] BIT NOT NULL DEFAULT (1),
        [IsActive] BIT NOT NULL DEFAULT (1),
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_EmailConfigurations] PRIMARY KEY ([Id] ASC)
    );
    PRINT '  -> Table [EmailConfigurations] created.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EmailTemplates')
BEGIN
    CREATE TABLE [dbo].[EmailTemplates] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [TemplateCode] NVARCHAR(50) NOT NULL,
        [Name] NVARCHAR(150) NOT NULL,
        [Category] NVARCHAR(50) NOT NULL DEFAULT ('QuotationFollowUp'),
        [SubjectTemplate] NVARCHAR(255) NOT NULL,
        [BodyTemplate] NVARCHAR(MAX) NOT NULL,
        [AvailablePlaceholders] NVARCHAR(500) NULL,
        [IsActive] BIT NOT NULL DEFAULT (1),
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT [PK_EmailTemplates] PRIMARY KEY ([Id] ASC)
    );
    PRINT '  -> Table [EmailTemplates] created.';
END
GO

-- ====================================================================================
-- 11. INITIAL SEED DATA (IDEMPOTENT)
-- ====================================================================================
PRINT 'Checking seed data...';

-- Seed Minute Types
IF NOT EXISTS (SELECT 1 FROM [dbo].[MinuteTypes] WHERE [Name] = 'Purchase Approval')
    INSERT INTO [dbo].[MinuteTypes] ([Name], [IsActive]) VALUES ('Purchase Approval', 1);

IF NOT EXISTS (SELECT 1 FROM [dbo].[MinuteTypes] WHERE [Name] = 'Stock Requisition')
    INSERT INTO [dbo].[MinuteTypes] ([Name], [IsActive]) VALUES ('Stock Requisition', 1);

IF NOT EXISTS (SELECT 1 FROM [dbo].[MinuteTypes] WHERE [Name] = 'HR Leave / Request')
    INSERT INTO [dbo].[MinuteTypes] ([Name], [IsActive]) VALUES ('HR Leave / Request', 1);

IF NOT EXISTS (SELECT 1 FROM [dbo].[MinuteTypes] WHERE [Name] = 'Financial Approval')
    INSERT INTO [dbo].[MinuteTypes] ([Name], [IsActive]) VALUES ('Financial Approval', 1);

IF NOT EXISTS (SELECT 1 FROM [dbo].[MinuteTypes] WHERE [Name] = 'General Memo / Policy')
    INSERT INTO [dbo].[MinuteTypes] ([Name], [IsActive]) VALUES ('General Memo / Policy', 1);

-- Seed Default General Channel
IF NOT EXISTS (SELECT 1 FROM [dbo].[Channels] WHERE [Name] = 'General')
    INSERT INTO [dbo].[Channels] ([Name], [Description], [DepartmentId], [IsPrivate], [IsActive], [CreatedBy])
    VALUES ('General', 'Company-wide general discussions channel', NULL, 0, 1, 'System');

PRINT 'IntraOffice Module Schema Migration completed successfully.';
GO
