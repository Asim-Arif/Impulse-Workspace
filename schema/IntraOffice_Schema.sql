-- ==============================================================================================
-- IntraOffice Communication Suite - Clean Table Creation Script for SMBI_AWM
-- Target: Microsoft SQL Server 2012 / 2014 / 2016 / 2019 / 2022
-- Integrates directly with existing [Departments] (deptid varchar(50)) and [Users] (UserName / EmpID)
-- ==============================================================================================

USE [SMBI_AWM];
GO

-- 1. CHANNELS (Departmental & Custom Group Chat)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Channels')
BEGIN
    CREATE TABLE [dbo].[Channels] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Name] NVARCHAR(200) NOT NULL,
        [Description] NVARCHAR(1000) NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [IsPrivate] BIT NOT NULL DEFAULT 0,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedBy] NVARCHAR(100) NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_Channels_DepartmentId] ON [dbo].[Channels]([DepartmentId]);
    CREATE NONCLUSTERED INDEX [IX_Channels_CreatedBy] ON [dbo].[Channels]([CreatedBy]);
    PRINT 'Created table: Channels';
END
GO

-- 2. CHANNEL MEMBERS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ChannelMembers')
BEGIN
    CREATE TABLE [dbo].[ChannelMembers] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ChannelId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [IsAdmin] BIT NOT NULL DEFAULT 0,
        [JoinedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_ChannelMembers_Channels] FOREIGN KEY ([ChannelId]) REFERENCES [dbo].[Channels]([Id]) ON DELETE CASCADE
    );
    CREATE UNIQUE NONCLUSTERED INDEX [IX_ChannelMembers_ChannelId_UserId] ON [dbo].[ChannelMembers]([ChannelId], [UserId]);
    CREATE NONCLUSTERED INDEX [IX_ChannelMembers_UserId] ON [dbo].[ChannelMembers]([UserId]);
    PRINT 'Created table: ChannelMembers';
END
GO

-- 3. MESSAGES (Direct 1-on-1 & Channel Messages)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Messages')
BEGIN
    CREATE TABLE [dbo].[Messages] (
        [Id] BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ChannelId] INT NULL,
        [SenderId] NVARCHAR(100) NOT NULL,
        [ReceiverId] NVARCHAR(100) NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [MessageType] INT NOT NULL DEFAULT 0, -- 0=Text, 1=File, 2=Audio/Voice, 3=System
        [IsRead] BIT NOT NULL DEFAULT 0,
        [IsDeleted] BIT NOT NULL DEFAULT 0,
        [ParentMessageId] BIGINT NULL,
        [SentAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [EditedAt] DATETIME2 NULL,
        CONSTRAINT [FK_Messages_Channels] FOREIGN KEY ([ChannelId]) REFERENCES [dbo].[Channels]([Id]) ON DELETE SET NULL,
        CONSTRAINT [FK_Messages_ParentMessage] FOREIGN KEY ([ParentMessageId]) REFERENCES [dbo].[Messages]([Id])
    );
    CREATE NONCLUSTERED INDEX [IX_Messages_ChannelId_SentAt] ON [dbo].[Messages]([ChannelId], [SentAt] DESC);
    CREATE NONCLUSTERED INDEX [IX_Messages_Sender_Receiver] ON [dbo].[Messages]([SenderId], [ReceiverId], [SentAt] DESC);
    CREATE NONCLUSTERED INDEX [IX_Messages_ParentMessageId] ON [dbo].[Messages]([ParentMessageId]);
    PRINT 'Created table: Messages';
END
GO

-- 4. MESSAGE ATTACHMENTS (Images, Documents, Voice Notes)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MessageAttachments')
BEGIN
    CREATE TABLE [dbo].[MessageAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [MessageId] BIGINT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT 0,
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_MessageAttachments_Messages] FOREIGN KEY ([MessageId]) REFERENCES [dbo].[Messages]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MessageAttachments_MessageId] ON [dbo].[MessageAttachments]([MessageId]);
    PRINT 'Created table: MessageAttachments';
END
GO

-- 5. ANNOUNCEMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Announcements')
BEGIN
    CREATE TABLE [dbo].[Announcements] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Title] NVARCHAR(400) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [Priority] INT NOT NULL DEFAULT 0, -- 0=Normal, 1=Important, 2=Critical
        [CreatedBy] NVARCHAR(100) NOT NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [IsPinned] BIT NOT NULL DEFAULT 0,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [IsAcknowledged] BIT NOT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [UpdatedAt] DATETIME2 NULL,
        [ExpiresAt] DATETIME2 NULL
    );
    CREATE NONCLUSTERED INDEX [IX_Announcements_CreatedBy] ON [dbo].[Announcements]([CreatedBy]);
    CREATE NONCLUSTERED INDEX [IX_Announcements_DepartmentId] ON [dbo].[Announcements]([DepartmentId]);
    CREATE NONCLUSTERED INDEX [IX_Announcements_CreatedAt] ON [dbo].[Announcements]([CreatedAt] DESC);
    PRINT 'Created table: Announcements';
END
GO

-- 6. ANNOUNCEMENT ATTACHMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AnnouncementAttachments')
BEGIN
    CREATE TABLE [dbo].[AnnouncementAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [AnnouncementId] INT NOT NULL,
        [FileName] NVARCHAR(MAX) NOT NULL,
        [FilePath] NVARCHAR(MAX) NOT NULL,
        [FileSize] BIGINT NOT NULL,
        [ContentType] NVARCHAR(MAX) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_AnnouncementAttachments_Announcements] FOREIGN KEY ([AnnouncementId]) REFERENCES [dbo].[Announcements]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_AnnouncementAttachments_AnnouncementId] ON [dbo].[AnnouncementAttachments]([AnnouncementId]);
    PRINT 'Created table: AnnouncementAttachments';
END
GO

-- 7. ANNOUNCEMENT ACKNOWLEDGMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AnnouncementAcknowledgments')
BEGIN
    CREATE TABLE [dbo].[AnnouncementAcknowledgments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [AnnouncementId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [AcknowledgedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_AnnouncementAcknowledgments_Announcements] FOREIGN KEY ([AnnouncementId]) REFERENCES [dbo].[Announcements]([Id]) ON DELETE CASCADE
    );
    CREATE UNIQUE NONCLUSTERED INDEX [IX_AnnouncementAcknowledgments_Announcement_User] ON [dbo].[AnnouncementAcknowledgments]([AnnouncementId], [UserId]);
    PRINT 'Created table: AnnouncementAcknowledgments';
END
GO

-- 8. TASK ITEMS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TaskItems')
BEGIN
    CREATE TABLE [dbo].[TaskItems] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Title] NVARCHAR(400) NOT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [AssignedTo] NVARCHAR(100) NULL,
        [AssignedBy] NVARCHAR(100) NOT NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [Priority] INT NOT NULL DEFAULT 1, -- 0=Low, 1=Normal, 2=High, 3=Urgent
        [Status] INT NOT NULL DEFAULT 0,   -- 0=Pending, 1=InProgress, 2=Completed, 3=Cancelled
        [DueDate] DATETIME2 NULL,
        [WhatsAppMessageSent] BIT NOT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [UpdatedAt] DATETIME2 NULL,
        [CompletedAt] DATETIME2 NULL
    );
    CREATE NONCLUSTERED INDEX [IX_TaskItems_AssignedTo] ON [dbo].[TaskItems]([AssignedTo]);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_AssignedBy] ON [dbo].[TaskItems]([AssignedBy]);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_Status] ON [dbo].[TaskItems]([Status]);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_DepartmentId] ON [dbo].[TaskItems]([DepartmentId]);
    PRINT 'Created table: TaskItems';
END
GO

-- 9. TASK COMMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TaskComments')
BEGIN
    CREATE TABLE [dbo].[TaskComments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [TaskId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_TaskComments_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_TaskComments_TaskId] ON [dbo].[TaskComments]([TaskId]);
    PRINT 'Created table: TaskComments';
END
GO

-- 10. TASK ATTACHMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TaskAttachments')
BEGIN
    CREATE TABLE [dbo].[TaskAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [TaskId] INT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT 0,
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_TaskAttachments_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_TaskAttachments_TaskId] ON [dbo].[TaskAttachments]([TaskId]);
    PRINT 'Created table: TaskAttachments';
END
GO

-- 11. USER PRESENCE (Real-time Online Status via SignalR)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'UserPresences')
BEGIN
    CREATE TABLE [dbo].[UserPresences] (
        [UserId] NVARCHAR(100) NOT NULL PRIMARY KEY,
        [Status] INT NOT NULL DEFAULT 0, -- 0=Offline, 1=Online, 2=Away, 3=Busy
        [LastSeen] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [ConnectionId] NVARCHAR(400) NULL
    );
    PRINT 'Created table: UserPresences';
END
GO

-- 12. MEETINGS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Meetings')
BEGIN
    CREATE TABLE [dbo].[Meetings] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Title] NVARCHAR(400) NOT NULL,
        [OrganizerId] NVARCHAR(100) NOT NULL,
        [MeetingUrl] NVARCHAR(1000) NULL,
        [ScheduledStartTime] DATETIME2 NOT NULL,
        [ScheduledEndTime] DATETIME2 NULL,
        [Status] INT NOT NULL DEFAULT 0, -- 0=Scheduled, 1=InProgress, 2=Completed, 3=Cancelled
        [MeetingMinutes] NVARCHAR(MAX) NULL,
        [IsReminderSent] BIT NOT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_Meetings_OrganizerId] ON [dbo].[Meetings]([OrganizerId]);
    CREATE NONCLUSTERED INDEX [IX_Meetings_ScheduledStartTime] ON [dbo].[Meetings]([ScheduledStartTime]);
    PRINT 'Created table: Meetings';
END
GO

-- 13. MEETING PARTICIPANTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MeetingParticipants')
BEGIN
    CREATE TABLE [dbo].[MeetingParticipants] (
        [MeetingId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [HasAttended] BIT NOT NULL DEFAULT 0,
        CONSTRAINT [PK_MeetingParticipants] PRIMARY KEY ([MeetingId], [UserId]),
        CONSTRAINT [FK_MeetingParticipants_Meetings] FOREIGN KEY ([MeetingId]) REFERENCES [dbo].[Meetings]([Id]) ON DELETE CASCADE
    );
    PRINT 'Created table: MeetingParticipants';
END
GO

-- 14. MINUTE TYPES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MinuteTypes')
BEGIN
    CREATE TABLE [dbo].[MinuteTypes] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Name] NVARCHAR(200) NOT NULL,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    PRINT 'Created table: MinuteTypes';
END
GO

-- 15. MINUTE APPROVALS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MinuteApprovals')
BEGIN
    CREATE TABLE [dbo].[MinuteApprovals] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Date] DATETIME2 NOT NULL,
        [No] NVARCHAR(100) NOT NULL,
        [Type] NVARCHAR(100) NOT NULL,
        [Subject] NVARCHAR(MAX) NOT NULL,
        [Points] NVARCHAR(MAX) NOT NULL,
        [ForwardToUserId] NVARCHAR(100) NULL,
        [Currency] NVARCHAR(20) NULL,
        [TotalAmount] DECIMAL(18,2) NULL,
        [AdvancePercentage] DECIMAL(5,2) NULL,
        [AdvanceAmount] DECIMAL(18,2) NULL,
        [IsUrgent] BIT NOT NULL DEFAULT 0,
        [CloseByInitiator] BIT NOT NULL DEFAULT 0,
        [CreatedByUserId] NVARCHAR(100) NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [Status] NVARCHAR(100) NOT NULL DEFAULT 'Pending',
        [SignaturePath] NVARCHAR(MAX) NULL,
        [PurchaseOrderValue] DECIMAL(18,2) NULL,
        [PurchaseAdvanceRecommend] DECIMAL(18,2) NULL,
        [PurchaseApprovedAmount] DECIMAL(18,2) NULL,
        [RequestedStockQty] DECIMAL(18,2) NULL,
        [CurrentStockQty] DECIMAL(18,2) NULL,
        [ApprovedStockQty] DECIMAL(18,2) NULL,
        [HRLeaveType] NVARCHAR(510) NULL,
        [FinancialType] NVARCHAR(510) NULL
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_CreatedByUserId] ON [dbo].[MinuteApprovals]([CreatedByUserId]);
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_ForwardToUserId] ON [dbo].[MinuteApprovals]([ForwardToUserId]);
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_Date] ON [dbo].[MinuteApprovals]([Date] DESC);
    PRINT 'Created table: MinuteApprovals';
END
GO

-- 16. MINUTE ATTACHMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MinuteAttachments')
BEGIN
    CREATE TABLE [dbo].[MinuteAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [MinuteApprovalId] INT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT 0,
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_MinuteAttachments_MinuteApprovals] FOREIGN KEY ([MinuteApprovalId]) REFERENCES [dbo].[MinuteApprovals]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteAttachments_MinuteApprovalId] ON [dbo].[MinuteAttachments]([MinuteApprovalId]);
    PRINT 'Created table: MinuteAttachments';
END
GO

-- 17. MINUTE WORKFLOW HISTORIES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MinuteWorkflowHistories')
BEGIN
    CREATE TABLE [dbo].[MinuteWorkflowHistories] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [MinuteApprovalId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [ActionTaken] NVARCHAR(100) NOT NULL,
        [Remarks] NVARCHAR(MAX) NULL,
        [SignaturePath] NVARCHAR(MAX) NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_MinuteWorkflowHistories_MinuteApprovals] FOREIGN KEY ([MinuteApprovalId]) REFERENCES [dbo].[MinuteApprovals]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteWorkflowHistories_MinuteApprovalId] ON [dbo].[MinuteWorkflowHistories]([MinuteApprovalId]);
    PRINT 'Created table: MinuteWorkflowHistories';
END
GO

-- 18. WORKFLOW TEMPLATES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'WorkflowTemplates')
BEGIN
    CREATE TABLE [dbo].[WorkflowTemplates] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Name] NVARCHAR(400) NOT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [RequestType] NVARCHAR(200) NOT NULL,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [UpdatedAt] DATETIME2 NULL
    );
    PRINT 'Created table: WorkflowTemplates';
END
GO

-- 19. WORKFLOW STEPS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'WorkflowSteps')
BEGIN
    CREATE TABLE [dbo].[WorkflowSteps] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [WorkflowTemplateId] INT NOT NULL,
        [StepOrder] INT NOT NULL,
        [StepName] NVARCHAR(MAX) NOT NULL,
        [ApproverType] INT NOT NULL, -- 0=Manager, 1=Role, 2=SpecificUser
        [SpecificUserId] NVARCHAR(100) NULL,
        [RoleName] NVARCHAR(MAX) NULL,
        [RequiredAction] INT NOT NULL,
        [SLAHours] INT NOT NULL DEFAULT 24,
        CONSTRAINT [FK_WorkflowSteps_WorkflowTemplates] FOREIGN KEY ([WorkflowTemplateId]) REFERENCES [dbo].[WorkflowTemplates]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_WorkflowSteps_WorkflowTemplateId] ON [dbo].[WorkflowSteps]([WorkflowTemplateId]);
    PRINT 'Created table: WorkflowSteps';
END
GO

-- 20. APPROVAL REQUESTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ApprovalRequests')
BEGIN
    CREATE TABLE [dbo].[ApprovalRequests] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [RequestNumber] NVARCHAR(100) NOT NULL,
        [Title] NVARCHAR(MAX) NOT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [WorkflowTemplateId] INT NOT NULL,
        [RequesterId] NVARCHAR(100) NOT NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [Priority] INT NOT NULL DEFAULT 1,
        [Status] INT NOT NULL DEFAULT 0,
        [RequestedDate] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [EffectiveDate] DATETIME2 NULL,
        [Amount] DECIMAL(18,2) NULL,
        [CostCenter] NVARCHAR(MAX) NULL,
        [Project] NVARCHAR(MAX) NULL,
        [CurrentStepOrder] INT NOT NULL DEFAULT 1,
        CONSTRAINT [FK_ApprovalRequests_WorkflowTemplates] FOREIGN KEY ([WorkflowTemplateId]) REFERENCES [dbo].[WorkflowTemplates]([Id])
    );
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_RequesterId] ON [dbo].[ApprovalRequests]([RequesterId]);
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_WorkflowTemplateId] ON [dbo].[ApprovalRequests]([WorkflowTemplateId]);
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_DepartmentId] ON [dbo].[ApprovalRequests]([DepartmentId]);
    PRINT 'Created table: ApprovalRequests';
END
GO

-- 21. APPROVAL REQUEST ATTACHMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ApprovalRequestAttachments')
BEGIN
    CREATE TABLE [dbo].[ApprovalRequestAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ApprovalRequestId] INT NOT NULL,
        [FileName] NVARCHAR(MAX) NOT NULL,
        [FilePath] NVARCHAR(MAX) NOT NULL,
        [FileSize] BIGINT NOT NULL,
        [ContentType] NVARCHAR(MAX) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_ApprovalRequestAttachments_ApprovalRequests] FOREIGN KEY ([ApprovalRequestId]) REFERENCES [dbo].[ApprovalRequests]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequestAttachments_ApprovalRequestId] ON [dbo].[ApprovalRequestAttachments]([ApprovalRequestId]);
    PRINT 'Created table: ApprovalRequestAttachments';
END
GO

-- 22. APPROVAL REQUEST HISTORIES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ApprovalRequestHistories')
BEGIN
    CREATE TABLE [dbo].[ApprovalRequestHistories] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ApprovalRequestId] INT NOT NULL,
        [StepOrder] INT NOT NULL,
        [ActionTaken] INT NOT NULL,
        [ActionById] NVARCHAR(100) NOT NULL,
        [ActionDate] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [Comments] NVARCHAR(MAX) NULL,
        CONSTRAINT [FK_ApprovalRequestHistories_ApprovalRequests] FOREIGN KEY ([ApprovalRequestId]) REFERENCES [dbo].[ApprovalRequests]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequestHistories_ApprovalRequestId] ON [dbo].[ApprovalRequestHistories]([ApprovalRequestId]);
    PRINT 'Created table: ApprovalRequestHistories';
END
GO

-- 23. APPROVAL DELEGATIONS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ApprovalDelegations')
BEGIN
    CREATE TABLE [dbo].[ApprovalDelegations] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [DelegatorId] NVARCHAR(100) NOT NULL,
        [DelegateeId] NVARCHAR(100) NOT NULL,
        [StartDate] DATETIME2 NOT NULL,
        [EndDate] DATETIME2 NOT NULL,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_ApprovalDelegations_DelegatorId] ON [dbo].[ApprovalDelegations]([DelegatorId]);
    CREATE NONCLUSTERED INDEX [IX_ApprovalDelegations_DelegateeId] ON [dbo].[ApprovalDelegations]([DelegateeId]);
    PRINT 'Created table: ApprovalDelegations';
END
GO

-- 24. STICKY NOTES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'StickyNotes')
BEGIN
    CREATE TABLE [dbo].[StickyNotes] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [UserId] NVARCHAR(100) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [Color] NVARCHAR(100) NOT NULL DEFAULT '#fffa65',
        [XPos] INT NOT NULL DEFAULT 100,
        [YPos] INT NOT NULL DEFAULT 100,
        [ReminderTime] DATETIME2 NULL,
        [IsReminderSent] BIT NOT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_StickyNotes_UserId] ON [dbo].[StickyNotes]([UserId]);
    PRINT 'Created table: StickyNotes';
END
GO

-- 25. GAME SCORES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'GameScores')
BEGIN
    CREATE TABLE [dbo].[GameScores] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [UserId] NVARCHAR(100) NOT NULL,
        [GameName] NVARCHAR(200) NOT NULL,
        [Score] INT NOT NULL,
        [PlayedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_GameScores_UserId] ON [dbo].[GameScores]([UserId]);
    PRINT 'Created table: GameScores';
END
GO

-- 26. DEFAULT SEED DATA (Minute Types)
IF NOT EXISTS (SELECT * FROM [dbo].[MinuteTypes] WHERE [Name] = 'General Memo')
BEGIN
    INSERT INTO [dbo].[MinuteTypes] ([Name]) VALUES 
    ('General Memo'),
    ('Purchase Approval'),
    ('Leave Request'),
    ('Financial Clearance'),
    ('Store & Inventory Request');
    PRINT 'Seeded default MinuteTypes';
END
GO

PRINT '==============================================================================================';
PRINT 'IntraOffice clean tables created successfully on SMBI_AWM!';
PRINT '==============================================================================================';
GO
