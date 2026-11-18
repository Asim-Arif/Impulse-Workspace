-- =============================================================================
-- Auto-generated Idempotent Schema Synchronization Script
-- Source Schema: Impulse_AWM_Web_Metro  ->  Target Schema: Impulse_AWM (Database-Independent)
-- Generated At: 2026-11-17 21:26:15
-- =============================================================================
SET NOCOUNT ON;
GO

-- =============================================================================
-- PHASE 1: USER-DEFINED TABLE TYPES
-- =============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.table_types tt JOIN sys.schemas s ON tt.schema_id = s.schema_id WHERE s.name = N'dbo' AND tt.name = N'Vouchers_TT')
BEGIN
    CREATE TYPE [dbo].[Vouchers_TT] AS TABLE (
        [SNo] [float] NULL,
        [VDate] [datetime] NULL,
        [VchrNo] [varchar](50) NOT NULL,
        [Accno] [nvarchar](255) NOT NULL,
        [Description] [varchar](8000) NULL,
        [Debit] [float] NULL,
        [Credit] [float] NULL,
        [balance] [float] NULL,
        [DpstSlip] [varchar](50) NULL,
        [CSNo] [bigint] NULL
    );
END
GO

-- =============================================================================
-- PHASE 2: NEW TABLES
-- =============================================================================
IF OBJECT_ID(N'[dbo].[`]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[`] (
        [Dimension1] [float] NULL,
        [Dimension2] [float] NULL,
        [Dimension3] [float] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[AnnouncementAcknowledgments]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[AnnouncementAcknowledgments] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [AnnouncementId] [int] NOT NULL,
        [UserId] [nvarchar](100) NOT NULL,
        [AcknowledgedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__Announcem__Ackno__328EB900] DEFAULT (getutcdate()),
        CONSTRAINT [PK__Announce__3214EC07B5766424] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[AnnouncementAcknowledgments]') AND name = N'IX_AnnouncementAcknowledgments_Announcement_User')
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX [IX_AnnouncementAcknowledgments_Announcement_User] ON [dbo].[AnnouncementAcknowledgments] ([AnnouncementId] ASC, [UserId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[AnnouncementAttachments]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[AnnouncementAttachments] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [AnnouncementId] [int] NOT NULL,
        [FileName] [nvarchar](MAX) NOT NULL,
        [FilePath] [nvarchar](MAX) NOT NULL,
        [FileSize] [bigint] NOT NULL,
        [ContentType] [nvarchar](MAX) NULL,
        [UploadedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__Announcem__Uploa__2EBE281C] DEFAULT (getutcdate()),
        CONSTRAINT [PK__Announce__3214EC07CB8CFD41] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[AnnouncementAttachments]') AND name = N'IX_AnnouncementAttachments_AnnouncementId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_AnnouncementAttachments_AnnouncementId] ON [dbo].[AnnouncementAttachments] ([AnnouncementId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[Announcements]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Announcements] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [Title] [nvarchar](400) NOT NULL,
        [Content] [nvarchar](MAX) NOT NULL,
        [Priority] [int] NOT NULL CONSTRAINT [DF__Announcem__Prior__28112A8D] DEFAULT ((0)),
        [CreatedBy] [nvarchar](100) NOT NULL,
        [DepartmentId] [varchar](50) NULL,
        [IsPinned] [bit] NOT NULL CONSTRAINT [DF__Announcem__IsPin__29054EC6] DEFAULT ((0)),
        [IsActive] [bit] NOT NULL CONSTRAINT [DF__Announcem__IsAct__29F972FF] DEFAULT ((1)),
        [IsAcknowledged] [bit] NOT NULL CONSTRAINT [DF__Announcem__IsAck__2AED9738] DEFAULT ((0)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__Announcem__Creat__2BE1BB71] DEFAULT (getutcdate()),
        [UpdatedAt] [datetime2](7) NULL,
        [ExpiresAt] [datetime2](7) NULL,
        CONSTRAINT [PK__Announce__3214EC071C9E4A96] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Announcements]') AND name = N'IX_Announcements_CreatedAt')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Announcements_CreatedAt] ON [dbo].[Announcements] ([CreatedAt] DESC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Announcements]') AND name = N'IX_Announcements_CreatedBy')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Announcements_CreatedBy] ON [dbo].[Announcements] ([CreatedBy] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Announcements]') AND name = N'IX_Announcements_DepartmentId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Announcements_DepartmentId] ON [dbo].[Announcements] ([DepartmentId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[AppNotifications]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[AppNotifications] (
        [Id] [nvarchar](50) NOT NULL,
        [Category] [int] NOT NULL CONSTRAINT [DF__AppNotifi__Categ__24FFB3B8] DEFAULT ((5)),
        [Title] [nvarchar](250) NOT NULL,
        [Message] [nvarchar](MAX) NOT NULL,
        [SenderName] [nvarchar](100) NULL,
        [TargetUserId] [nvarchar](100) NULL,
        [ActionUrl] [nvarchar](500) NOT NULL CONSTRAINT [DF__AppNotifi__Actio__25F3D7F1] DEFAULT ('/'),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__AppNotifi__Creat__26E7FC2A] DEFAULT (getutcdate()),
        [IsRead] [bit] NOT NULL CONSTRAINT [DF__AppNotifi__IsRea__27DC2063] DEFAULT ((0)),
        [IsReadReceipt] [bit] NOT NULL CONSTRAINT [DF__AppNotifi__IsRea__28D0449C] DEFAULT ((0)),
        CONSTRAINT [PK__AppNotif__3214EC07F94A087F] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[AppNotifications]') AND name = N'IX_AppNotifications_TargetUser_IsRead')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_AppNotifications_TargetUser_IsRead] ON [dbo].[AppNotifications] ([TargetUserId] ASC, [IsRead] ASC, [CreatedAt] DESC);
END
GO
IF OBJECT_ID(N'[dbo].[ApprovalDelegations]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ApprovalDelegations] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [DelegatorId] [nvarchar](100) NOT NULL,
        [DelegateeId] [nvarchar](100) NOT NULL,
        [StartDate] [datetime2](7) NOT NULL,
        [EndDate] [datetime2](7) NOT NULL,
        [IsActive] [bit] NOT NULL CONSTRAINT [DF__ApprovalD__IsAct__7921317A] DEFAULT ((1)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__ApprovalD__Creat__7A1555B3] DEFAULT (getutcdate()),
        CONSTRAINT [PK__Approval__3214EC07E3B2B4FC] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ApprovalDelegations]') AND name = N'IX_ApprovalDelegations_DelegateeId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ApprovalDelegations_DelegateeId] ON [dbo].[ApprovalDelegations] ([DelegateeId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ApprovalDelegations]') AND name = N'IX_ApprovalDelegations_DelegatorId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ApprovalDelegations_DelegatorId] ON [dbo].[ApprovalDelegations] ([DelegatorId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[ApprovalRequestAttachments]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ApprovalRequestAttachments] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [ApprovalRequestId] [int] NOT NULL,
        [FileName] [nvarchar](MAX) NOT NULL,
        [FilePath] [nvarchar](MAX) NOT NULL,
        [FileSize] [bigint] NOT NULL,
        [ContentType] [nvarchar](MAX) NULL,
        [UploadedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__ApprovalR__Uploa__71800FB2] DEFAULT (getutcdate()),
        CONSTRAINT [PK__Approval__3214EC07A42762C5] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ApprovalRequestAttachments]') AND name = N'IX_ApprovalRequestAttachments_ApprovalRequestId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequestAttachments_ApprovalRequestId] ON [dbo].[ApprovalRequestAttachments] ([ApprovalRequestId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[ApprovalRequestHistories]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ApprovalRequestHistories] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [ApprovalRequestId] [int] NOT NULL,
        [StepOrder] [int] NOT NULL,
        [ActionTaken] [int] NOT NULL,
        [ActionById] [nvarchar](100) NOT NULL,
        [ActionDate] [datetime2](7) NOT NULL CONSTRAINT [DF__ApprovalR__Actio__7550A096] DEFAULT (getutcdate()),
        [Comments] [nvarchar](MAX) NULL,
        CONSTRAINT [PK__Approval__3214EC07F6BA9FE3] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ApprovalRequestHistories]') AND name = N'IX_ApprovalRequestHistories_ApprovalRequestId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequestHistories_ApprovalRequestId] ON [dbo].[ApprovalRequestHistories] ([ApprovalRequestId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[ApprovalRequests]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ApprovalRequests] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [RequestNumber] [nvarchar](100) NOT NULL,
        [Title] [nvarchar](MAX) NOT NULL,
        [Description] [nvarchar](MAX) NULL,
        [WorkflowTemplateId] [int] NOT NULL,
        [RequesterId] [nvarchar](100) NOT NULL,
        [DepartmentId] [varchar](50) NULL,
        [Priority] [int] NOT NULL CONSTRAINT [DF__ApprovalR__Prior__6AD31223] DEFAULT ((1)),
        [Status] [int] NOT NULL CONSTRAINT [DF__ApprovalR__Statu__6BC7365C] DEFAULT ((0)),
        [RequestedDate] [datetime2](7) NOT NULL CONSTRAINT [DF__ApprovalR__Reque__6CBB5A95] DEFAULT (getutcdate()),
        [EffectiveDate] [datetime2](7) NULL,
        [Amount] [decimal](18, 2) NULL,
        [CostCenter] [nvarchar](MAX) NULL,
        [Project] [nvarchar](MAX) NULL,
        [CurrentStepOrder] [int] NOT NULL CONSTRAINT [DF__ApprovalR__Curre__6DAF7ECE] DEFAULT ((1)),
        CONSTRAINT [PK__Approval__3214EC07F72624C1] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ApprovalRequests]') AND name = N'IX_ApprovalRequests_DepartmentId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_DepartmentId] ON [dbo].[ApprovalRequests] ([DepartmentId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ApprovalRequests]') AND name = N'IX_ApprovalRequests_RequesterId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_RequesterId] ON [dbo].[ApprovalRequests] ([RequesterId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ApprovalRequests]') AND name = N'IX_ApprovalRequests_WorkflowTemplateId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_WorkflowTemplateId] ON [dbo].[ApprovalRequests] ([WorkflowTemplateId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[BudgetingAccounts]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[BudgetingAccounts] (
        [AccNo] [nvarchar](255) NOT NULL,
        [AccTitle] [nvarchar](255) NULL,
        [SubAccOf] [nvarchar](50) NULL,
        [Type] [nvarchar](50) NULL,
        [opendate] [nvarchar](50) NULL,
        [openbal] [money] NULL,
        [Balance] [money] NULL,
        [Parent] [bit] NULL CONSTRAINT [DF__Budgeting__Paren__79763735] DEFAULT ((0)),
        [Active] [bit] NULL CONSTRAINT [DF__Budgeting__Activ__7A6A5B6E] DEFAULT ((0)),
        [CreateDT] [datetime] NULL CONSTRAINT [DF__Budgeting__Creat__7B5E7FA7] DEFAULT (getdate()),
        [UserName] [varchar](50) NULL,
        [ComputerName] [varchar](50) NULL,
        [Revision] [real] NULL,
        CONSTRAINT [PK_BudgetingAccounts] PRIMARY KEY CLUSTERED ([AccNo] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[ChannelMembers]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ChannelMembers] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [ChannelId] [int] NOT NULL,
        [UserId] [nvarchar](100) NOT NULL,
        [IsAdmin] [bit] NOT NULL CONSTRAINT [DF__ChannelMe__IsAdm__16E69E8B] DEFAULT ((0)),
        [JoinedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__ChannelMe__Joine__17DAC2C4] DEFAULT (getutcdate()),
        CONSTRAINT [PK__ChannelM__3214EC07A549B85D] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ChannelMembers]') AND name = N'IX_ChannelMembers_ChannelId_UserId')
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX [IX_ChannelMembers_ChannelId_UserId] ON [dbo].[ChannelMembers] ([ChannelId] ASC, [UserId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ChannelMembers]') AND name = N'IX_ChannelMembers_UserId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ChannelMembers_UserId] ON [dbo].[ChannelMembers] ([UserId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[Channels]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Channels] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [Name] [nvarchar](200) NOT NULL,
        [Description] [nvarchar](1000) NULL,
        [DepartmentId] [varchar](50) NULL,
        [IsPrivate] [bit] NOT NULL CONSTRAINT [DF__Channels__IsPriv__1221E96E] DEFAULT ((0)),
        [IsActive] [bit] NOT NULL CONSTRAINT [DF__Channels__IsActi__13160DA7] DEFAULT ((1)),
        [CreatedBy] [nvarchar](100) NOT NULL,
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__Channels__Create__140A31E0] DEFAULT (getutcdate()),
        CONSTRAINT [PK__Channels__3214EC070613CF0B] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Channels]') AND name = N'IX_Channels_CreatedBy')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Channels_CreatedBy] ON [dbo].[Channels] ([CreatedBy] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Channels]') AND name = N'IX_Channels_DepartmentId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Channels_DepartmentId] ON [dbo].[Channels] ([DepartmentId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[ChequePrintHistory]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ChequePrintHistory] (
        [Sno] [bigint] NULL,
        [Payee] [varchar](500) NULL,
        [PrintDT] [datetime] NULL,
        [Designation] [varchar](255) NULL,
        [Company] [varchar](255) NULL,
        [EntryDT] [datetime] NULL CONSTRAINT [DF_ChequePrintHistory_EntryDT] DEFAULT (getdate()),
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        CONSTRAINT [PK__ChequePr__F57BD2D74EC42028] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[Courier_Issuance]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Courier_Issuance] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [DT] [datetime] NULL,
        [TrackingNo] [varchar](50) NULL,
        [Description] [varchar](255) NULL,
        [VendorName] [varchar](50) NULL,
        [VendorAddress] [varchar](255) NULL,
        [IssuedBy] [varchar](50) NULL,
        [ConcernedPerson] [varchar](50) NULL,
        [ReceivedBy] [varchar](50) NULL,
        [Self] [bit] NULL CONSTRAINT [DF_Courier_Issuance_Self] DEFAULT ((1)),
        [Paid] [bit] NULL CONSTRAINT [DF_Courier_Issuance_Paid] DEFAULT ((0)),
        [VendID] [varchar](50) NULL,
        [Weight] [varchar](50) NULL,
        [Dimension] [varchar](50) NULL,
        [CourierStatus] [bit] NULL CONSTRAINT [DF_Courier_Issuance_CourierStatus] DEFAULT ((0)),
        [MachineName] [varchar](50) NULL,
        [UserName] [varchar](50) NULL,
        [Parcel_Qty] [real] NULL,
        [Vehicle_No] [varchar](50) NULL,
        CONSTRAINT [PK_Courier_Issuance] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[EmailTemplates]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[EmailTemplates] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [TemplateCode] [nvarchar](50) NOT NULL,
        [Name] [nvarchar](150) NOT NULL,
        [Category] [nvarchar](50) NOT NULL CONSTRAINT [DF__EmailTemp__Categ__3906AC65] DEFAULT ('QuotationFollowUp'),
        [SubjectTemplate] [nvarchar](255) NOT NULL,
        [BodyTemplate] [nvarchar](MAX) NOT NULL,
        [AvailablePlaceholders] [nvarchar](500) NULL,
        [IsActive] [bit] NOT NULL CONSTRAINT [DF__EmailTemp__IsAct__39FAD09E] DEFAULT ((1)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__EmailTemp__Creat__3AEEF4D7] DEFAULT (getutcdate()),
        CONSTRAINT [PK_EmailTemplates] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[EmpGatePass]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[EmpGatePass] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [EmpID] [varchar](50) NULL,
        [DT] [datetime] NULL,
        [TimeOut] [char](5) NULL,
        [TimeIn] [char](5) NULL,
        [EmpID_Recommendation] [varchar](50) NULL,
        [EmpID_Approval] [varchar](50) NULL,
        [Purpose] [varchar](1000) NULL,
        [GPType] [tinyint] NULL,
        [TimeTotal] [char](5) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__EmpGatePa__DTEnt__2184288F] DEFAULT (getdate()),
        [EmpTimes_RefID] [int] NULL,
        [HP_TimeOut] [char](5) NULL,
        [HP_TimeIn] [char](5) NULL,
        [EmpTimes_In_RefID] [int] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[EmpGatePass_ApprovalAuthorities]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[EmpGatePass_ApprovalAuthorities] (
        [EmpID] [varchar](50) NOT NULL,
        CONSTRAINT [PK_EmpGatePass_ApprovalAuthorities] PRIMARY KEY CLUSTERED ([EmpID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[EntityAttachments]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[EntityAttachments] (
        [Id] [bigint] IDENTITY(1, 1) NOT NULL,
        [EntityType] [varchar](50) NOT NULL,
        [EntityId] [varchar](100) NOT NULL,
        [EntitySecondaryId] [varchar](100) NULL,
        [OriginalFileName] [nvarchar](500) NOT NULL,
        [StoredFileName] [nvarchar](500) NOT NULL,
        [FilePath] [nvarchar](1000) NOT NULL,
        [FileSize] [bigint] NOT NULL CONSTRAINT [DF__EntityAtt__FileS__1276EF29] DEFAULT ((0)),
        [FileExtension] [varchar](20) NOT NULL,
        [ContentType] [varchar](150) NULL,
        [Remarks] [nvarchar](500) NULL,
        [UploadedBy] [nvarchar](100) NOT NULL,
        [UploadedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__EntityAtt__Uploa__136B1362] DEFAULT (getdate()),
        [IsDeleted] [bit] NOT NULL CONSTRAINT [DF__EntityAtt__IsDel__145F379B] DEFAULT ((0)),
        [DeletedBy] [nvarchar](100) NULL,
        [DeletedAt] [datetime2](7) NULL,
        CONSTRAINT [PK__EntityAt__3214EC07C3B1679F] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[EntityAttachments]') AND name = N'IX_EntityAttachments_Lookup')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_EntityAttachments_Lookup] ON [dbo].[EntityAttachments] ([EntityType] ASC, [EntityId] ASC, [EntitySecondaryId] ASC, [IsDeleted] ASC) INCLUDE ([OriginalFileName], [FileSize], [FileExtension], [UploadedAt], [UploadedBy]);
END
GO
IF OBJECT_ID(N'[dbo].[EntityNotes]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[EntityNotes] (
        [Id] [bigint] IDENTITY(1, 1) NOT NULL,
        [EntityType] [varchar](50) NOT NULL,
        [EntityId] [varchar](100) NOT NULL,
        [EntitySecondaryId] [varchar](100) NULL,
        [NoteText] [nvarchar](MAX) NOT NULL CONSTRAINT [DF__EntityNot__NoteT__173BA446] DEFAULT (''),
        [LastModifiedBy] [nvarchar](100) NOT NULL,
        [LastModifiedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__EntityNot__LastM__182FC87F] DEFAULT (getdate()),
        CONSTRAINT [PK__EntityNo__3214EC07A3BA6B28] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[EntityNotes]') AND name = N'IX_EntityNotes_Lookup')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_EntityNotes_Lookup] ON [dbo].[EntityNotes] ([EntityType] ASC, [EntityId] ASC, [EntitySecondaryId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[EnumValues]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[EnumValues] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [EnumName] [varchar](50) NOT NULL,
        [EnumValue] [int] NOT NULL,
        [EnumDescription] [varchar](50) NOT NULL,
        CONSTRAINT [PK_EnumValues] PRIMARY KEY CLUSTERED ([EnumName] ASC, [EnumValue] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[FCO_Qualities]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FCO_Qualities] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [Quality] [varchar](50) NOT NULL,
        CONSTRAINT [PK_FCO_Qualities] PRIMARY KEY CLUSTERED ([Quality] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[FinancialBudgeting]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FinancialBudgeting] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [FinancialYear] [varchar](50) NOT NULL,
        [AccNo] [nvarchar](255) NOT NULL,
        [Month_July] [bigint] NULL,
        [Month_August] [bigint] NULL,
        [Month_September] [bigint] NULL,
        [Month_October] [bigint] NULL,
        [Month_November] [bigint] NULL,
        [Month_December] [bigint] NULL,
        [Month_January] [bigint] NULL,
        [Month_February] [bigint] NULL,
        [Month_March] [bigint] NULL,
        [Month_April] [bigint] NULL,
        [Month_May] [bigint] NULL,
        [Month_June] [bigint] NULL,
        [Jul_Sep] [bigint] NULL,
        [Jul_Sep_Revision] [bigint] NULL,
        [Oct_Dec] [bigint] NULL,
        [Oct_Dec_Revision] [bigint] NULL,
        [Jan_Mar] [bigint] NULL,
        [Jan_Mar_Revision] [bigint] NULL,
        [Apr_Jun] [bigint] NULL,
        [Apr_Jun_Revision] [bigint] NULL,
        [Month_July_R] [bigint] NULL,
        [Month_August_R] [bigint] NULL,
        [Month_September_R] [bigint] NULL,
        [Month_October_R] [bigint] NULL,
        [Month_November_R] [bigint] NULL,
        [Month_December_R] [bigint] NULL,
        [Month_January_R] [bigint] NULL,
        [Month_February_R] [bigint] NULL,
        [Month_March_R] [bigint] NULL,
        [Month_April_R] [bigint] NULL,
        [Month_May_R] [bigint] NULL,
        [Month_June_R] [bigint] NULL,
        CONSTRAINT [PK_FinancialBudgeting] PRIMARY KEY CLUSTERED ([FinancialYear] ASC, [AccNo] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[FinancialBudgeting_Expenses]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FinancialBudgeting_Expenses] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [FinancialYear] [varchar](50) NOT NULL,
        [AccNo] [nvarchar](255) NOT NULL,
        [Proposed] [bigint] NULL,
        [Revised] [bigint] NULL,
        CONSTRAINT [PK_FinancialBudgeting_Expenses] PRIMARY KEY CLUSTERED ([FinancialYear] ASC, [AccNo] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Adjustment]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Adjustment] (
        [EntryID] [int] NOT NULL,
        [AssetID] [varchar](50) NULL,
        [DT] [datetime] NULL,
        [Remarks] [varchar](500) NULL,
        [Cost] [float] NULL,
        [Adjustment] [float] NULL,
        [Attachment_Type] [varchar](50) NULL,
        [PDFAttachments] [varchar](255) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Approval]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Approval] (
        [EntryID] [int] NOT NULL,
        [AssetID] [varchar](50) NULL,
        [ForwardBy] [varchar](50) NULL,
        [DTEntry] [datetime] NULL,
        [MachineName] [varchar](50) NULL,
        [ForwardTo] [varchar](50) NULL,
        [DTEntry_Done] [datetime] NULL,
        [MachineName_Done] [varchar](50) NULL,
        [Remarks] [varchar](5000) NULL,
        [Approval] [int] NULL,
        [ForwardBy_Designation] [varchar](255) NULL,
        [ForwardTo_Designation] [varchar](255) NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Depreciation_Detail]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Depreciation_Detail] (
        [EntryID] [int] NOT NULL,
        [FA_D_M_RefID] [int] NULL,
        [AssetID] [varchar](50) NULL,
        [Cost] [float] NULL,
        [BookValue] [float] NULL,
        [DisposedValue] [real] NULL,
        [OpeningBalance] [float] NULL,
        [AccountingYear] [float] NULL,
        [AccumulatedBalance] [float] NULL,
        [WrittenDownValue] [float] NULL,
        [Rate] [real] NULL,
        [Cost_Last] [real] NULL,
        [DisposedValue_Last] [real] NULL,
        [AcquisitionCost_Last] [real] NULL,
        [Adjustment_Opening] [real] NULL,
        [AdjustedValue] [real] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Depreciation_Detail_Monthwise]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Depreciation_Detail_Monthwise] (
        [EntryID] [int] NOT NULL,
        [FA_D_M_RefID] [int] NULL,
        [AssetID] [varchar](50) NULL,
        [Cost] [real] NULL,
        [BookValue] [real] NULL,
        [DisposedValue] [real] NULL,
        [OpeningBalance] [real] NULL,
        [AccountingYear] [real] NULL,
        [AccumulatedBalance] [real] NULL,
        [WrittenDownValue] [real] NULL,
        [Rate] [real] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Depreciation_Main]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Depreciation_Main] (
        [EntryID] [int] NOT NULL,
        [DTFrom] [datetime] NULL,
        [DTTo] [datetime] NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Depreciation_Main_Monthwise]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Depreciation_Main_Monthwise] (
        [EntryID] [int] NOT NULL,
        [DTFrom] [datetime] NULL,
        [DTTo] [datetime] NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Disposals]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Disposals] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [AssetID] [varchar](50) NOT NULL,
        [DisposalDate] [datetime] NOT NULL,
        [DisposalType] [varchar](50) NULL,
        [DisposalAmount] [decimal](18, 2) NOT NULL CONSTRAINT [DF__FixedAsse__Dispo__07C4568C] DEFAULT ((0)),
        [Buyer] [varchar](255) NULL,
        [Remarks] [varchar](MAX) NULL,
        [UserName] [varchar](100) NULL,
        [MachineName] [varchar](100) NULL,
        [DTEntry] [datetime] NOT NULL CONSTRAINT [DF__FixedAsse__DTEnt__08B87AC5] DEFAULT (getdate()),
        CONSTRAINT [PK_FixedAssets_Disposals] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Head_Changes]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Head_Changes] (
        [EntryID] [int] NOT NULL,
        [AssetID] [varchar](50) NULL,
        [FA_M_RefID_Old] [int] NULL,
        [FA_M_RefID_New] [int] NULL,
        [Remarks] [varchar](255) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Master_DepreciationRate]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Master_DepreciationRate] (
        [EntryID] [int] NOT NULL,
        [FA_M_RefID] [int] NULL,
        [DTFrom] [datetime] NULL,
        [DTTo] [datetime] NULL,
        [PercentRate] [real] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[FixedAssets_Master_Groups]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Master_Groups] (
        [Group_Name] [varchar](255) NOT NULL,
        [FA_M_RefID] [int] NOT NULL,
        [Group_Name_Main] [varchar](50) NOT NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[ForeignCustomers_InvoiceTo]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ForeignCustomers_InvoiceTo] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [CustCode] [varchar](50) NULL,
        [Country] [varchar](50) NULL,
        [InvoiceTo] [varchar](1000) NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[GameScores]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GameScores] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [UserId] [nvarchar](100) NOT NULL,
        [GameName] [nvarchar](200) NOT NULL,
        [Score] [int] NOT NULL,
        [PlayedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__GameScore__Playe__039EBFED] DEFAULT (getutcdate()),
        CONSTRAINT [PK__GameScor__3214EC07B3468C5C] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[GameScores]') AND name = N'IX_GameScores_UserId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_GameScores_UserId] ON [dbo].[GameScores] ([UserId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[GateConcernedContacts]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GateConcernedContacts] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [EmpID] [varchar](50) NOT NULL,
        [ContactName] [varchar](150) NULL,
        [EmailAddress] [varchar](150) NULL,
        [WhatsAppNo] [varchar](50) NULL,
        [Active] [bit] NOT NULL CONSTRAINT [DF__GateConce__Activ__0E71541B] DEFAULT ((1)),
        [LastUpdatedBy] [varchar](50) NULL,
        [LastUpdatedAt] [datetime] NOT NULL CONSTRAINT [DF__GateConce__LastU__0F657854] DEFAULT (getdate()),
        CONSTRAINT [PK_GateConcernedContacts] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[GateNotificationLog]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GateNotificationLog] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [SourceModule] [varchar](50) NOT NULL,
        [SourceEntryId] [int] NOT NULL,
        [ChannelName] [varchar](20) NOT NULL,
        [RecipientName] [varchar](150) NULL,
        [RecipientAddress] [varchar](150) NULL,
        [MessageBody] [varchar](MAX) NULL,
        [DeliveryStatus] [varchar](30) NOT NULL CONSTRAINT [DF__GateNotif__Deliv__1241E4FF] DEFAULT ('Queued'),
        [DeliveryResponse] [varchar](500) NULL,
        [CreatedAt] [datetime] NOT NULL CONSTRAINT [DF__GateNotif__Creat__13360938] DEFAULT (getdate()),
        [SentAt] [datetime] NULL,
        [CreatedBy] [varchar](50) NULL,
        CONSTRAINT [PK_GateNotificationLog] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[GatePass_Return]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GatePass_Return] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [GatePassNo] [varchar](50) NULL,
        [DT] [datetime] NULL,
        [SenderName] [varchar](255) NULL,
        [PersonName] [varchar](255) NULL,
        [VehicleNo] [varchar](50) NULL,
        [Description] [varchar](255) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__GatePass___DTEnt__18EEE28E] DEFAULT (getdate()),
        [Returnable] [bit] NULL CONSTRAINT [DF__GatePass___Retur__19E306C7] DEFAULT ((0)),
        [IssdTo] [varchar](255) NULL,
        [GatePassType] [varchar](50) NULL,
        [ReturnDT] [datetime] NULL,
        [DeptID_From] [varchar](50) NULL,
        [DeptID_To] [varchar](50) NULL,
        [ApprovedBy_Name] [varchar](50) NULL,
        [ApprovedBy_Designation] [varchar](50) NULL,
        [VendID] [int] NULL,
        [VM_AccNo] [varchar](50) NULL,
        CONSTRAINT [PK_GatePass_Return] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[GatePass_Return_Details]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GatePass_Return_Details] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [MEGO_RefID] [int] NULL,
        [Detail] [varchar](255) NULL,
        [Unit] [varchar](50) NULL,
        [Packing] [varchar](255) NULL,
        [Qty] [int] NULL,
        CONSTRAINT [PK_GatePass_Return_Details] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[GatePass_Return_Receiving]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GatePass_Return_Receiving] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__GatePass___DTEnt__1EA7BBE4] DEFAULT (getdate()),
        [ME_GO_RefID] [int] NULL,
        [RcvDT] [datetime] NULL,
        CONSTRAINT [PK_GatePass_Return_Receiving] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[GatePass_Return_Receiving_Detail]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GatePass_Return_Receiving_Detail] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [ME_GO_Rcv_RefID] [int] NULL,
        [ME_GO_D_RefID] [int] NULL,
        [QtyRcvd] [int] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[GatePassTypes]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GatePassTypes] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [GatePassType] [varchar](50) NOT NULL,
        [Active] [bit] NOT NULL CONSTRAINT [DF__GatePassT__Activ__161275E3] DEFAULT ((1)),
        CONSTRAINT [PK_GatePassTypes] PRIMARY KEY CLUSTERED ([GatePassType] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[Hub_Names]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Hub_Names] (
        [Hub_Name] [varchar](50) NOT NULL,
        CONSTRAINT [PK_Hub_Names] PRIMARY KEY CLUSTERED ([Hub_Name] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[Items_Parts]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Items_Parts] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [ItemID] [varchar](50) NOT NULL,
        [ItemID_Part] [varchar](50) NOT NULL,
        [Qty] [real] NULL,
        [ProcessID] [int] NULL,
        CONSTRAINT [PK_Items_Parts] PRIMARY KEY CLUSTERED ([ItemID] ASC, [ItemID_Part] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[License_Keys]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[License_Keys] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [LicenseKey] [varchar](4000) NULL,
        [MachineName] [varchar](50) NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[MaterialInward]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MaterialInward] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [PoNumber] [varchar](50) NULL,
        [SupplierName] [varchar](150) NULL,
        [VehicleEntryId] [int] NULL,
        [VehicleNumber] [varchar](50) NULL,
        [DriverName] [varchar](100) NULL,
        [DeliveryChallanNo] [varchar](50) NULL,
        [InvoiceNo] [varchar](50) NULL,
        [Weighment] [decimal](18, 3) NULL,
        [EntryTime] [datetime] NOT NULL CONSTRAINT [DF__MaterialI__Entry__2648DDAC] DEFAULT (getdate()),
        [Status] [varchar](20) NOT NULL CONSTRAINT [DF__MaterialI__Statu__273D01E5] DEFAULT ('Open'),
        [CreatedBy] [varchar](50) NULL,
        [UpdatedBy] [varchar](50) NULL,
        [CreatedAt] [datetime] NOT NULL CONSTRAINT [DF__MaterialI__Creat__2831261E] DEFAULT (getdate()),
        [UpdatedAt] [datetime] NULL,
        [Remarks] [varchar](255) NULL,
        CONSTRAINT [PK_MaterialInward] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[MaterialInwardLines]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MaterialInwardLines] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [MaterialInwardId] [int] NOT NULL,
        [MaterialName] [varchar](150) NOT NULL,
        [Quantity] [decimal](18, 3) NOT NULL CONSTRAINT [DF__MaterialI__Quant__2B0D92C9] DEFAULT ((0)),
        [Uom] [varchar](25) NULL,
        CONSTRAINT [PK_MaterialInwardLines] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[MaterialOutward]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MaterialOutward] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [SalesOrderNo] [varchar](50) NULL,
        [DeliveryOrderNo] [varchar](50) NULL,
        [CustomerName] [varchar](150) NULL,
        [VehicleEntryId] [int] NULL,
        [VehicleNumber] [varchar](50) NULL,
        [DriverName] [varchar](100) NULL,
        [Weighment] [decimal](18, 3) NULL,
        [EntryTime] [datetime] NOT NULL CONSTRAINT [DF__MaterialO__Entry__2DE9FF74] DEFAULT (getdate()),
        [Status] [varchar](20) NOT NULL CONSTRAINT [DF__MaterialO__Statu__2EDE23AD] DEFAULT ('Dispatched'),
        [CreatedBy] [varchar](50) NULL,
        [UpdatedBy] [varchar](50) NULL,
        [CreatedAt] [datetime] NOT NULL CONSTRAINT [DF__MaterialO__Creat__2FD247E6] DEFAULT (getdate()),
        [UpdatedAt] [datetime] NULL,
        [Remarks] [varchar](255) NULL,
        CONSTRAINT [PK_MaterialOutward] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[MaterialOutwardLines]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MaterialOutwardLines] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [MaterialOutwardId] [int] NOT NULL,
        [MaterialName] [varchar](150) NOT NULL,
        [Quantity] [decimal](18, 3) NOT NULL CONSTRAINT [DF__MaterialO__Quant__32AEB491] DEFAULT ((0)),
        [Uom] [varchar](25) NULL,
        CONSTRAINT [PK_MaterialOutwardLines] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[MeetingParticipants]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MeetingParticipants] (
        [MeetingId] [int] NOT NULL,
        [UserId] [nvarchar](100) NOT NULL,
        [HasAttended] [bit] NOT NULL CONSTRAINT [DF__MeetingPa__HasAt__4D42AF3C] DEFAULT ((0)),
        CONSTRAINT [PK_MeetingParticipants] PRIMARY KEY CLUSTERED ([MeetingId] ASC, [UserId] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[Meetings]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Meetings] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [Title] [nvarchar](400) NOT NULL,
        [OrganizerId] [nvarchar](100) NOT NULL,
        [MeetingUrl] [nvarchar](1000) NULL,
        [ScheduledStartTime] [datetime2](7) NOT NULL,
        [ScheduledEndTime] [datetime2](7) NULL,
        [Status] [int] NOT NULL CONSTRAINT [DF__Meetings__Status__487DFA1F] DEFAULT ((0)),
        [MeetingMinutes] [nvarchar](MAX) NULL,
        [IsReminderSent] [bit] NOT NULL CONSTRAINT [DF__Meetings__IsRemi__49721E58] DEFAULT ((0)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__Meetings__Create__4A664291] DEFAULT (getutcdate()),
        [VoiceNotePath] [nvarchar](500) NULL,
        CONSTRAINT [PK__Meetings__3214EC077DD341EA] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Meetings]') AND name = N'IX_Meetings_OrganizerId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Meetings_OrganizerId] ON [dbo].[Meetings] ([OrganizerId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Meetings]') AND name = N'IX_Meetings_ScheduledStartTime')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Meetings_ScheduledStartTime] ON [dbo].[Meetings] ([ScheduledStartTime] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[MessageAttachments]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MessageAttachments] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [MessageId] [bigint] NOT NULL,
        [FileName] [nvarchar](510) NOT NULL,
        [FilePath] [nvarchar](1000) NOT NULL,
        [FileSize] [bigint] NOT NULL CONSTRAINT [DF__MessageAt__FileS__234C7570] DEFAULT ((0)),
        [ContentType] [nvarchar](200) NULL,
        [UploadedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__MessageAt__Uploa__244099A9] DEFAULT (getutcdate()),
        CONSTRAINT [PK__MessageA__3214EC078A3068C0] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MessageAttachments]') AND name = N'IX_MessageAttachments_MessageId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MessageAttachments_MessageId] ON [dbo].[MessageAttachments] ([MessageId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[Messages]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Messages] (
        [Id] [bigint] IDENTITY(1, 1) NOT NULL,
        [ChannelId] [int] NULL,
        [SenderId] [nvarchar](100) NOT NULL,
        [ReceiverId] [nvarchar](100) NULL,
        [Content] [nvarchar](MAX) NOT NULL,
        [MessageType] [int] NOT NULL CONSTRAINT [DF__Messages__Messag__1BAB53A8] DEFAULT ((0)),
        [IsRead] [bit] NOT NULL CONSTRAINT [DF__Messages__IsRead__1C9F77E1] DEFAULT ((0)),
        [IsDeleted] [bit] NOT NULL CONSTRAINT [DF__Messages__IsDele__1D939C1A] DEFAULT ((0)),
        [ParentMessageId] [bigint] NULL,
        [SentAt] [datetime2](7) NOT NULL CONSTRAINT [DF__Messages__SentAt__1E87C053] DEFAULT (getutcdate()),
        [EditedAt] [datetime2](7) NULL,
        CONSTRAINT [PK__Messages__3214EC07CBD60510] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Messages]') AND name = N'IX_Messages_ChannelId_SentAt')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Messages_ChannelId_SentAt] ON [dbo].[Messages] ([ChannelId] ASC, [SentAt] DESC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Messages]') AND name = N'IX_Messages_ParentMessageId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Messages_ParentMessageId] ON [dbo].[Messages] ([ParentMessageId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Messages]') AND name = N'IX_Messages_Sender_Receiver')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Messages_Sender_Receiver] ON [dbo].[Messages] ([SenderId] ASC, [ReceiverId] ASC, [SentAt] DESC);
END
GO
IF OBJECT_ID(N'[dbo].[MinuteApprovals]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MinuteApprovals] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [Date] [datetime2](7) NOT NULL,
        [No] [nvarchar](100) NOT NULL,
        [Type] [nvarchar](100) NOT NULL,
        [Subject] [nvarchar](MAX) NOT NULL,
        [Points] [nvarchar](MAX) NOT NULL,
        [ForwardToUserId] [nvarchar](100) NULL,
        [Currency] [nvarchar](20) NULL,
        [TotalAmount] [decimal](18, 2) NULL,
        [AdvancePercentage] [decimal](5, 2) NULL,
        [AdvanceAmount] [decimal](18, 2) NULL,
        [IsUrgent] [bit] NOT NULL CONSTRAINT [DF__MinuteApp__IsUrg__54E3D104] DEFAULT ((0)),
        [CloseByInitiator] [bit] NOT NULL CONSTRAINT [DF__MinuteApp__Close__55D7F53D] DEFAULT ((0)),
        [CreatedByUserId] [nvarchar](100) NOT NULL,
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__MinuteApp__Creat__56CC1976] DEFAULT (getutcdate()),
        [Status] [nvarchar](100) NOT NULL CONSTRAINT [DF__MinuteApp__Statu__57C03DAF] DEFAULT ('Pending'),
        [SignaturePath] [nvarchar](MAX) NULL,
        [PurchaseOrderValue] [decimal](18, 2) NULL,
        [PurchaseAdvanceRecommend] [decimal](18, 2) NULL,
        [PurchaseApprovedAmount] [decimal](18, 2) NULL,
        [RequestedStockQty] [decimal](18, 2) NULL,
        [CurrentStockQty] [decimal](18, 2) NULL,
        [ApprovedStockQty] [decimal](18, 2) NULL,
        [HRLeaveType] [nvarchar](510) NULL,
        [FinancialType] [nvarchar](510) NULL,
        [IsRead] [bit] NOT NULL CONSTRAINT [DF_MinuteApprovals_IsRead] DEFAULT ((0)),
        [ReadAt] [datetime2](7) NULL,
        CONSTRAINT [PK__MinuteAp__3214EC07779B52EC] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MinuteApprovals]') AND name = N'IX_MinuteApprovals_CreatedByUserId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_CreatedByUserId] ON [dbo].[MinuteApprovals] ([CreatedByUserId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MinuteApprovals]') AND name = N'IX_MinuteApprovals_Date')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_Date] ON [dbo].[MinuteApprovals] ([Date] DESC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MinuteApprovals]') AND name = N'IX_MinuteApprovals_ForwardToUserId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_ForwardToUserId] ON [dbo].[MinuteApprovals] ([ForwardToUserId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[MinuteAttachments]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MinuteAttachments] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [MinuteApprovalId] [int] NOT NULL,
        [FileName] [nvarchar](510) NOT NULL,
        [FilePath] [nvarchar](1000) NOT NULL,
        [FileSize] [bigint] NOT NULL CONSTRAINT [DF__MinuteAtt__FileS__5A9CAA5A] DEFAULT ((0)),
        [ContentType] [nvarchar](200) NULL,
        [UploadedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__MinuteAtt__Uploa__5B90CE93] DEFAULT (getutcdate()),
        CONSTRAINT [PK__MinuteAt__3214EC074513454C] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MinuteAttachments]') AND name = N'IX_MinuteAttachments_MinuteApprovalId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MinuteAttachments_MinuteApprovalId] ON [dbo].[MinuteAttachments] ([MinuteApprovalId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[MinuteTypes]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MinuteTypes] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [Name] [nvarchar](200) NOT NULL,
        [IsActive] [bit] NOT NULL CONSTRAINT [DF__MinuteTyp__IsAct__51134020] DEFAULT ((1)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__MinuteTyp__Creat__52076459] DEFAULT (getutcdate()),
        CONSTRAINT [PK__MinuteTy__3214EC0716C983E0] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[MinuteWorkflowHistories]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MinuteWorkflowHistories] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [MinuteApprovalId] [int] NOT NULL,
        [UserId] [nvarchar](100) NOT NULL,
        [ActionTaken] [nvarchar](100) NOT NULL,
        [Remarks] [nvarchar](MAX) NULL,
        [SignaturePath] [nvarchar](MAX) NULL,
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__MinuteWor__Creat__5F615F77] DEFAULT (getutcdate()),
        CONSTRAINT [PK__MinuteWo__3214EC07064DC571] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MinuteWorkflowHistories]') AND name = N'IX_MinuteWorkflowHistories_MinuteApprovalId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MinuteWorkflowHistories_MinuteApprovalId] ON [dbo].[MinuteWorkflowHistories] ([MinuteApprovalId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[MiscEntries_GateInTemp]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MiscEntries_GateInTemp] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [GatePassNo] [varchar](50) NULL,
        [DT] [datetime] NULL,
        [PersonName] [varchar](255) NULL,
        [VehicleNo] [varchar](50) NULL,
        [Description] [varchar](255) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__MiscEntri__DTEnt__358B213C] DEFAULT (getdate()),
        [EmpID] [varchar](50) NULL,
        [GatePassType] [int] NULL CONSTRAINT [DF__MiscEntri__GateP__367F4575] DEFAULT ((0)),
        [DeptID] [varchar](50) NULL,
        [Approved_HOD] [bit] NULL CONSTRAINT [DF__MiscEntri__Appro__377369AE] DEFAULT ((0)),
        [Approved] [bit] NULL CONSTRAINT [DF__MiscEntri__Appro__38678DE7] DEFAULT ((0)),
        [Approved_Finance] [bit] NULL CONSTRAINT [DF__MiscEntri__Appro__395BB220] DEFAULT ((0)),
        [GatePassType2] [int] NULL CONSTRAINT [DF__MiscEntri__GateP__3A4FD659] DEFAULT ((0)),
        [VendID] [int] NULL,
        [VendNameManual] [varchar](500) NULL,
        [EmpNameManual] [varchar](50) NULL,
        [GatePassTypeName] [varchar](50) NULL,
        [VM_AccNo] [varchar](50) NULL,
        CONSTRAINT [PK_MiscEntries_GateInTemp] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[MiscEntries_GateInTemp_Details]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[MiscEntries_GateInTemp_Details] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [MEGO_RefID] [int] NULL,
        [Detail] [varchar](255) NULL,
        [Unit] [varchar](50) NULL,
        [Rate] [real] NULL,
        [Qty] [real] NULL,
        [ReturnType] [varchar](20) NULL,
        [ReturnDT] [datetime] NULL,
        [QtyRcvd] [real] NULL,
        [RcvdDT] [datetime] NULL,
        [RcvdBy] [varchar](50) NULL,
        [RcvdRemarks] [varchar](255) NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[PPC_Order_Item_Hub_Schedules]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Hub_Schedules] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [OrderNo] [varchar](50) NOT NULL,
        [ItemID] [varchar](50) NOT NULL,
        [Hub_Name] [varchar](50) NOT NULL,
        [PlannedQty] [int] NOT NULL CONSTRAINT [DF__PPC_Order__Plann__7CBCB834] DEFAULT ((0)),
        [StartDate] [datetime] NOT NULL,
        [EndDate] [datetime] NOT NULL,
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__PPC_Order__Creat__7DB0DC6D] DEFAULT (getutcdate()),
        [CreatedBy] [nvarchar](100) NOT NULL,
        CONSTRAINT [PK__PPC_Orde__F57BD2D718232441] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[PPC_Order_Item_Planning]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Planning] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [OrderNo] [varchar](100) NOT NULL,
        [ItemID] [varchar](50) NOT NULL,
        [OrderQty] [int] NOT NULL,
        [StockQty] [int] NOT NULL CONSTRAINT [DF__PPC_Order__Stock__0369B5C3] DEFAULT ((0)),
        [TotalPurchaseQty] [int] NOT NULL CONSTRAINT [DF__PPC_Order__Total__045DD9FC] DEFAULT ((0)),
        [ProductionQty] [int] NOT NULL CONSTRAINT [DF__PPC_Order__Produ__0551FE35] DEFAULT ((0)),
        [Remarks] [nvarchar](250) NULL,
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__PPC_Order__Creat__0646226E] DEFAULT (getutcdate()),
        [CreatedBy] [nvarchar](100) NOT NULL,
        [SFStockQty] [int] NOT NULL CONSTRAINT [DF_PPC_SFStockQty] DEFAULT ((0)),
        [RejectedStockQty] [int] NOT NULL CONSTRAINT [DF_PPC_ItemPlan_RejectedStockQty] DEFAULT ((0)),
        [RejectedSFQty] [int] NOT NULL CONSTRAINT [DF_PPC_ItemPlan_RejectedSFQty] DEFAULT ((0)),
        CONSTRAINT [PK__PPC_Orde__F57BD2D728F1C969] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[PPC_Order_Item_Purchases]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Purchases] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [OrderNo] [varchar](50) NOT NULL,
        [ItemID] [varchar](50) NOT NULL,
        [ProcessID] [int] NOT NULL,
        [VendID] [int] NULL,
        [PurchaseQty] [int] NOT NULL,
        [PurchaseRate] [float] NULL CONSTRAINT [DF__PPC_Order__Purch__6F62BD16] DEFAULT ((0)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__PPC_Order__Creat__7056E14F] DEFAULT (getutcdate()),
        [CreatedBy] [nvarchar](100) NOT NULL,
        [VendIssued_RefID] [bigint] NULL,
        [MasterPONo] [varchar](50) NULL,
        [IssuedAt] [datetime2](7) NULL,
        [IssuedBy] [nvarchar](100) NULL,
        [ExcessQtyPercentage] [real] NOT NULL CONSTRAINT [DF__PPC_Order__Exces__7D7BD243] DEFAULT ((50)),
        [Status] [varchar](30) NOT NULL CONSTRAINT [DF_PPC_ItemPurch_Status] DEFAULT ('Active'),
        [OriginalPlannedQty] [int] NULL,
        [ShortClosedQty] [int] NOT NULL CONSTRAINT [DF_PPC_ItemPurch_ShortClosedQty] DEFAULT ((0)),
        [ExceptionNotes] [nvarchar](500) NULL,
        [UpdatedAt] [datetime2](7) NULL,
        [UpdatedBy] [nvarchar](100) NULL,
        CONSTRAINT [PK__PPC_Orde__F57BD2D7DB56D5CD] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[PPC_Order_Item_SF_Allocations]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_SF_Allocations] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [OrderNo] [varchar](100) NOT NULL,
        [ItemID] [varchar](50) NOT NULL,
        [SOO_RefID] [int] NOT NULL,
        [ProcessID] [int] NOT NULL,
        [ProcessName] [nvarchar](100) NULL,
        [LotNo] [varchar](100) NULL,
        [BatchNo] [varchar](100) NULL,
        [Shelf_RefID] [int] NULL,
        [StoreLocation] [nvarchar](200) NULL,
        [AllocatedQty] [int] NOT NULL,
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__PPC_Order__Creat__4BAF6C85] DEFAULT (getutcdate()),
        [CreatedBy] [nvarchar](100) NULL,
        CONSTRAINT [PK__PPC_Orde__F57BD2D7E090D1E6] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[PPC_Order_Item_SF_Allocations]') AND name = N'IX_PPC_SF_Alloc_Order_Item')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_PPC_SF_Alloc_Order_Item] ON [dbo].[PPC_Order_Item_SF_Allocations] ([OrderNo] ASC, [ItemID] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[PPC_Order_Planning_Exceptions]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Planning_Exceptions] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [OrderNo] [varchar](100) NOT NULL,
        [ItemID] [varchar](50) NOT NULL,
        [ExceptionType] [varchar](50) NOT NULL,
        [AffectedQty] [int] NOT NULL CONSTRAINT [DF__PPC_Order__Affec__7758E499] DEFAULT ((0)),
        [SourceRef] [varchar](100) NULL,
        [Reason] [nvarchar](500) NOT NULL,
        [CreatedBy] [nvarchar](100) NOT NULL,
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__PPC_Order__Creat__784D08D2] DEFAULT (getutcdate()),
        CONSTRAINT [PK__PPC_Orde__F57BD2D7A12C86BE] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[PPC_Order_Planning_Exceptions]') AND name = N'IX_PPC_Exceptions_Order_Item')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_PPC_Exceptions_Order_Item] ON [dbo].[PPC_Order_Planning_Exceptions] ([OrderNo] ASC, [ItemID] ASC, [CreatedAt] DESC);
END
GO
IF OBJECT_ID(N'[dbo].[PPC_Order_Planning_Master]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Planning_Master] (
        [OrderNo] [varchar](50) NOT NULL,
        [PlannedBy] [nvarchar](100) NOT NULL,
        [PlannedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__PPC_Order__Plann__760FBAA5] DEFAULT (getutcdate()),
        [Status] [varchar](20) NOT NULL CONSTRAINT [DF__PPC_Order__Statu__7703DEDE] DEFAULT ('Planned'),
        [Notes] [nvarchar](500) NULL,
        CONSTRAINT [PK__PPC_Orde__C3907C74627ED538] PRIMARY KEY CLUSTERED ([OrderNo] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[PrintEmpAbsentList_Summary]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[PrintEmpAbsentList_Summary] (
        [EmpID] [varchar](50) NULL,
        [Total_Absents] [int] NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[Processes_Purchase]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Processes_Purchase] (
        [ProcessID] [int] NOT NULL,
        CONSTRAINT [PK_Processes_Purchase] PRIMARY KEY CLUSTERED ([ProcessID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[ProcessGroup_Hub_Supervisors]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ProcessGroup_Hub_Supervisors] (
        [GroupID] [int] NOT NULL,
        [Hub_Name] [varchar](50) NOT NULL,
        [UserID] [int] NOT NULL,
        [UserName] [nvarchar](100) NOT NULL,
        [AssignedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__ProcessGr__Assig__13D527B6] DEFAULT (getdate()),
        CONSTRAINT [PK_ProcessGroup_Hub_Supervisors] PRIMARY KEY CLUSTERED ([GroupID] ASC, [Hub_Name] ASC, [UserID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[ProductionDeletionRequests]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ProductionDeletionRequests] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [RequestType] [varchar](50) NOT NULL,
        [EntityRefID] [bigint] NOT NULL,
        [LotNo] [varchar](50) NOT NULL,
        [OrderNo] [varchar](50) NULL,
        [ItemCode] [varchar](50) NULL,
        [ItemName] [nvarchar](200) NULL,
        [ProcessID] [int] NULL,
        [ProcessName] [nvarchar](100) NULL,
        [MakerID] [bigint] NULL,
        [MakerName] [nvarchar](150) NULL,
        [Qty] [numeric](18, 2) NOT NULL CONSTRAINT [DF__ProductionD__Qty__60DF9395] DEFAULT ((0)),
        [RequestedBy] [varchar](50) NOT NULL,
        [RequestedDT] [datetime] NOT NULL CONSTRAINT [DF__Productio__Reque__61D3B7CE] DEFAULT (getdate()),
        [Reason] [nvarchar](500) NOT NULL,
        [MachineName] [varchar](100) NULL,
        [Status] [varchar](20) NOT NULL CONSTRAINT [DF__Productio__Statu__62C7DC07] DEFAULT ('Pending'),
        [ReviewedBy] [varchar](50) NULL,
        [ReviewedDT] [datetime] NULL,
        [DirectorRemarks] [nvarchar](500) NULL,
        [TaskId] [int] NULL,
        CONSTRAINT [PK__Producti__3214EC077DAA6159] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ProductionDeletionRequests]') AND name = N'IX_ProdDelReq_Active')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ProdDelReq_Active] ON [dbo].[ProductionDeletionRequests] ([EntityRefID] ASC, [Status] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ProductionDeletionRequests]') AND name = N'IX_ProdDelReq_LotNo')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ProdDelReq_LotNo] ON [dbo].[ProductionDeletionRequests] ([LotNo] ASC, [Status] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ProductionDeletionRequests]') AND name = N'IX_ProdDelReq_TypeStatus')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ProdDelReq_TypeStatus] ON [dbo].[ProductionDeletionRequests] ([RequestType] ASC, [Status] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[PTC_Printing]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[PTC_Printing] (
        [OrderNo] [varchar](50) NULL,
        [ItemCode] [varchar](50) NULL,
        [Qty] [int] NULL,
        [CustCode] [varchar](50) NULL,
        [Country] [varchar](50) NULL,
        [DT] [datetime] NULL,
        [ItemName] [varchar](255) NULL,
        [ItemSize] [varchar](255) NULL,
        [SizeUnit] [varchar](50) NULL,
        [Type] [varchar](50) NULL,
        [Packaging] [varchar](50) NULL,
        [DeliveryDT] [datetime] NULL,
        [SNo] [int] NULL,
        [ProcessID] [int] NULL,
        [Description] [varchar](255) NULL,
        [RcvdQty] [int] NULL,
        [LotNo] [varchar](50) NULL,
        [VendID] [int] NULL,
        [VenderName] [varchar](255) NULL,
        [RcvDT] [datetime] NULL,
        [IssDT] [datetime] NULL,
        [MinProcessID] [int] NULL,
        [Wastage] [int] NULL,
        [PassQty] [int] NULL,
        [EmpID] [varchar](50) NULL,
        [Name] [varchar](255) NULL,
        [ReWorkQty] [int] NULL,
        [CompItemCode] [varchar](50) NULL,
        [Code] [varchar](50) NULL,
        [ProcessNameUrdu] [nvarchar](255) NULL,
        [ItemNameUrdu] [nvarchar](255) NULL,
        [RcvEmpName] [varchar](50) NULL,
        [Rate] [real] NULL,
        [MasterPONo] [varchar](50) NULL,
        [SteelBatchNo] [varchar](50) NULL,
        [ForgeBatchNo] [varchar](50) NULL,
        [SFRemarks] [varchar](255) NULL,
        [SplitQty] [int] NULL,
        [SplitType] [tinyint] NULL,
        [ToOrderNo] [varchar](50) NULL,
        [ToLotNo] [varchar](50) NULL,
        [FromOrderNo] [varchar](50) NULL,
        [FromLotNo] [varchar](50) NULL,
        [SplitQtyFrom] [int] NULL,
        [FromItemCode] [varchar](50) NULL,
        [ToItemCode] [varchar](50) NULL,
        [RecieptID] [varchar](50) NULL,
        [IssSFLotNo] [varchar](50) NULL,
        [IssSFRemarks] [varchar](50) NULL,
        [IssSFQty] [int] NULL,
        [LotLastEntryID] [int] NULL,
        [ShippedQty] [int] NULL,
        [FromLotPolisher] [varchar](50) NULL,
        [FromLotOriginator] [varchar](50) NULL,
        [RcvTemperValue] [varchar](50) NULL,
        [IssSpecialInstructions] [varchar](50) NULL,
        [InternalRefNo] [varchar](50) NULL,
        [PGReference] [varchar](50) NULL,
        [ReceivedEntryID] [int] NULL,
        [Employees] [varchar](255) NULL,
        [Quality] [varchar](255) NULL,
        [RepairLots] [varchar](255) NULL,
        [LotTransferRemarks] [varchar](255) NULL,
        [LotTransferRemarksFrom] [varchar](255) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [Hub_Name] [varchar](50) NULL,
        [VI_DT] [datetime] NULL,
        [VID_Batch_No] [varchar](50) NULL,
        [MasterPODT] [datetime] NULL,
        [StoreName] [varchar](50) NULL,
        [Location] [varchar](50) NULL,
        [Lots_List_Batch_No] [varchar](50) NULL,
        [Reference_LotNo] [varchar](50) NULL,
        [Mill_Certificate_No] [varchar](50) NULL,
        [Mill_Certificate_No_From_RM] [varchar](50) NULL,
        [HRC_From] [int] NULL,
        [HRC_To] [int] NULL,
        [Lots_List_Forge_Batch_No] [varchar](50) NULL,
        [FinQuality] [varchar](50) NULL,
        [SpecialInstructions_FCC] [varchar](50) NULL,
        [StampInstructions] [varchar](50) NULL,
        [PackingInstructions] [varchar](50) NULL,
        [MDMA] [varchar](50) NULL,
        [SFDA_Listing_No] [varchar](50) NULL,
        [MD_Group] [varchar](50) NULL,
        [ItemPic] [image] NULL,
        [FOI_Stamps] [varchar](50) NULL,
        [SpecialInstruction] [varchar](50) NULL,
        [BatchNo] [varchar](50) NULL,
        [Maker_Second_Name] [varchar](50) NULL,
        [Finishing] [varchar](50) NULL,
        [ItemOrderNo] [varchar](50) NULL,
        [DeliveryDate] [datetime] NULL,
        [ItemRemarks] [varchar](50) NULL,
        [VI_UserName] [varchar](50) NULL,
        [VI_MachineName] [varchar](50) NULL,
        [IssEmpID] [varchar](50) NULL,
        [Iss_Emp_Name] [varchar](50) NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[RawMaterialIssuanceDetail_Hidden_From_Maker_Billing]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[RawMaterialIssuanceDetail_Hidden_From_Maker_Billing] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [RMID_RefID] [int] NOT NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__RawMateri__DTEnt__334DD30F] DEFAULT (getdate())
    );
END
GO
IF OBJECT_ID(N'[dbo].[RemoteAccessRequests]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[RemoteAccessRequests] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [UserID] [int] NOT NULL,
        [UserName] [varchar](50) NOT NULL,
        [ClientIP] [varchar](50) NOT NULL,
        [RequestDT] [datetime] NOT NULL CONSTRAINT [DF__RemoteAcc__Reque__6FB7C2D1] DEFAULT (getdate()),
        [Reason] [nvarchar](255) NULL,
        [Status] [varchar](20) NOT NULL CONSTRAINT [DF__RemoteAcc__Statu__70ABE70A] DEFAULT ('Pending'),
        [DurationMinutes] [int] NULL,
        [ExpiresAt] [datetime] NULL,
        [ActionDT] [datetime] NULL,
        [ActionByUserID] [int] NULL,
        [ActionByUserName] [varchar](50) NULL,
        [Remarks] [nvarchar](255) NULL,
        CONSTRAINT [PK__RemoteAc__F57BD2D71F9CB8F2] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[RemoteAccessRequests]') AND name = N'IX_RemoteAccessRequests_Expires')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_RemoteAccessRequests_Expires] ON [dbo].[RemoteAccessRequests] ([Status] ASC, [ExpiresAt] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[RemoteAccessRequests]') AND name = N'IX_RemoteAccessRequests_User_Status')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_RemoteAccessRequests_User_Status] ON [dbo].[RemoteAccessRequests] ([UserID] ASC, [Status] ASC, [RequestDT] DESC);
END
GO
IF OBJECT_ID(N'[dbo].[RM_PO_Payment_Terms]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[RM_PO_Payment_Terms] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [PaymentTerms] [varchar](50) NOT NULL,
        CONSTRAINT [PK_RM_PO_Payment_Terms] PRIMARY KEY CLUSTERED ([PaymentTerms] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[StickyNotes]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[StickyNotes] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [UserId] [nvarchar](100) NOT NULL,
        [Content] [nvarchar](MAX) NOT NULL,
        [Color] [nvarchar](100) NOT NULL CONSTRAINT [DF__StickyNot__Color__7CF1C25E] DEFAULT ('#fffa65'),
        [XPos] [int] NOT NULL CONSTRAINT [DF__StickyNote__XPos__7DE5E697] DEFAULT ((100)),
        [YPos] [int] NOT NULL CONSTRAINT [DF__StickyNote__YPos__7EDA0AD0] DEFAULT ((100)),
        [ReminderTime] [datetime2](7) NULL,
        [IsReminderSent] [bit] NOT NULL CONSTRAINT [DF__StickyNot__IsRem__7FCE2F09] DEFAULT ((0)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__StickyNot__Creat__00C25342] DEFAULT (getutcdate()),
        CONSTRAINT [PK__StickyNo__3214EC078F3A9043] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[StickyNotes]') AND name = N'IX_StickyNotes_UserId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_StickyNotes_UserId] ON [dbo].[StickyNotes] ([UserId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[SystemFieldConfigurations]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[SystemFieldConfigurations] (
        [SettingID] [int] IDENTITY(1, 1) NOT NULL,
        [Module] [varchar](50) NOT NULL,
        [FormKey] [varchar](50) NOT NULL,
        [FieldKey] [varchar](50) NOT NULL,
        [IsRequired] [bit] NOT NULL CONSTRAINT [DF__SystemFie__IsReq__58D45D79] DEFAULT ((1)),
        [IsVisible] [bit] NOT NULL CONSTRAINT [DF__SystemFie__IsVis__59C881B2] DEFAULT ((1)),
        [Description] [varchar](255) NULL,
        CONSTRAINT [PK__SystemFi__54372AFD0EC556FD] PRIMARY KEY CLUSTERED ([SettingID] ASC),
        CONSTRAINT [UQ_SystemFieldConfigurations] UNIQUE NONCLUSTERED ([Module] ASC, [FormKey] ASC, [FieldKey] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[Task_Assignees]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Task_Assignees] (
        [TaskID] [int] NOT NULL,
        [UserID] [int] NOT NULL,
        [UserName] [nvarchar](100) NOT NULL,
        [AssignedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__Task_Assi__Assig__1D5E91F0] DEFAULT (getdate()),
        CONSTRAINT [PK_Task_Assignees] PRIMARY KEY CLUSTERED ([TaskID] ASC, [UserID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[Task_Roles]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Task_Roles] (
        [TaskID] [int] NOT NULL,
        [RoleName] [nvarchar](50) NOT NULL,
        [AssignedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__Task_Role__Assig__198E010C] DEFAULT (getdate()),
        CONSTRAINT [PK_Task_Roles] PRIMARY KEY CLUSTERED ([TaskID] ASC, [RoleName] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[TaskAttachments]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[TaskAttachments] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [TaskId] [int] NOT NULL,
        [FileName] [nvarchar](510) NOT NULL,
        [FilePath] [nvarchar](1000) NOT NULL,
        [FileSize] [bigint] NOT NULL CONSTRAINT [DF__TaskAttac__FileS__3FE8B41E] DEFAULT ((0)),
        [ContentType] [nvarchar](200) NULL,
        [UploadedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__TaskAttac__Uploa__40DCD857] DEFAULT (getutcdate()),
        CONSTRAINT [PK__TaskAtta__3214EC0751AAD9D4] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[TaskAttachments]') AND name = N'IX_TaskAttachments_TaskId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TaskAttachments_TaskId] ON [dbo].[TaskAttachments] ([TaskId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[TaskComments]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[TaskComments] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [TaskId] [int] NOT NULL,
        [UserId] [nvarchar](100) NOT NULL,
        [Content] [nvarchar](MAX) NOT NULL,
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__TaskComme__Creat__3C18233A] DEFAULT (getutcdate()),
        CONSTRAINT [PK__TaskComm__3214EC07C3FE3E22] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[TaskComments]') AND name = N'IX_TaskComments_TaskId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TaskComments_TaskId] ON [dbo].[TaskComments] ([TaskId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[TaskItems]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[TaskItems] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [Title] [nvarchar](400) NOT NULL,
        [Description] [nvarchar](MAX) NULL,
        [AssignedTo] [nvarchar](100) NULL,
        [AssignedBy] [nvarchar](100) NOT NULL,
        [DepartmentId] [varchar](50) NULL,
        [Priority] [int] NOT NULL CONSTRAINT [DF__TaskItems__Prior__365F49E4] DEFAULT ((1)),
        [Status] [int] NOT NULL CONSTRAINT [DF__TaskItems__Statu__37536E1D] DEFAULT ((0)),
        [DueDate] [datetime2](7) NULL,
        [WhatsAppMessageSent] [bit] NOT NULL CONSTRAINT [DF__TaskItems__Whats__38479256] DEFAULT ((0)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__TaskItems__Creat__393BB68F] DEFAULT (getutcdate()),
        [UpdatedAt] [datetime2](7) NULL,
        [CompletedAt] [datetime2](7) NULL,
        [SourceEntityType] [nvarchar](50) NULL,
        [SourceEntityRefId] [nvarchar](100) NULL,
        [TargetRole] [nvarchar](50) NULL,
        [CompletedBy] [nvarchar](100) NULL,
        [ActionUrl] [nvarchar](300) NULL,
        [DueWarningSent] [bit] NOT NULL CONSTRAINT [DF_TaskItems_DueWarningSent] DEFAULT ((0)),
        [OverdueWarningSent] [bit] NOT NULL CONSTRAINT [DF_TaskItems_OverdueWarningSent] DEFAULT ((0)),
        [AdditionalAssigneeIds] [nvarchar](MAX) NULL,
        [AssignedToNames] [nvarchar](MAX) NULL,
        [EmailMessageSent] [bit] NOT NULL CONSTRAINT [DF_TaskItems_EmailSent] DEFAULT ((0)),
        [IsRead] [bit] NOT NULL CONSTRAINT [DF_TaskItems_IsRead] DEFAULT ((0)),
        [StartedAt] [datetime2](7) NULL,
        CONSTRAINT [PK__TaskItem__3214EC0731100EC0] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[TaskItems]') AND name = N'IX_TaskItems_AssignedBy')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TaskItems_AssignedBy] ON [dbo].[TaskItems] ([AssignedBy] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[TaskItems]') AND name = N'IX_TaskItems_AssignedTo')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TaskItems_AssignedTo] ON [dbo].[TaskItems] ([AssignedTo] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[TaskItems]') AND name = N'IX_TaskItems_DepartmentId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TaskItems_DepartmentId] ON [dbo].[TaskItems] ([DepartmentId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[TaskItems]') AND name = N'IX_TaskItems_SourceEntity')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TaskItems_SourceEntity] ON [dbo].[TaskItems] ([SourceEntityType] ASC, [SourceEntityRefId] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[TaskItems]') AND name = N'IX_TaskItems_Status')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TaskItems_Status] ON [dbo].[TaskItems] ([Status] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[TaskItems]') AND name = N'IX_TaskItems_TargetRole_Status')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TaskItems_TargetRole_Status] ON [dbo].[TaskItems] ([TargetRole] ASC, [Status] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[User_Roles]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[User_Roles] (
        [User_Role] [varchar](50) NOT NULL,
        CONSTRAINT [PK_User_Roles] PRIMARY KEY CLUSTERED ([User_Role] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[UserPresences]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[UserPresences] (
        [UserId] [nvarchar](100) NOT NULL,
        [Status] [int] NOT NULL CONSTRAINT [DF__UserPrese__Statu__44AD693B] DEFAULT ((0)),
        [LastSeen] [datetime2](7) NOT NULL CONSTRAINT [DF__UserPrese__LastS__45A18D74] DEFAULT (getutcdate()),
        [ConnectionId] [nvarchar](400) NULL,
        CONSTRAINT [PK__UserPres__1788CC4CE3BE2363] PRIMARY KEY CLUSTERED ([UserId] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[UserPushSubscriptions]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[UserPushSubscriptions] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [UserID] [int] NULL,
        [UserName] [nvarchar](100) NOT NULL,
        [Endpoint] [nvarchar](1000) NOT NULL,
        [P256dh] [nvarchar](255) NOT NULL,
        [Auth] [nvarchar](255) NOT NULL,
        [DeviceName] [varchar](500) NULL,
        [CreatedAt] [datetime] NOT NULL CONSTRAINT [DF__UserPushS__Creat__45377920] DEFAULT (getdate()),
        [LastUsedAt] [datetime] NULL,
        CONSTRAINT [PK__UserPush__3214EC07A243A25B] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[Users_User_Roles]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Users_User_Roles] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [UserID] [int] NOT NULL,
        [User_Role] [varchar](50) NOT NULL,
        CONSTRAINT [PK__Users_Us__F57BD2D7A3950047] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[VehicleCategories]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[VehicleCategories] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [CategoryName] [varchar](50) NOT NULL,
        [SortOrder] [int] NOT NULL CONSTRAINT [DF__VehicleCa__SortO__3E20673D] DEFAULT ((0)),
        [Active] [bit] NOT NULL CONSTRAINT [DF__VehicleCa__Activ__3F148B76] DEFAULT ((1)),
        CONSTRAINT [PK_VehicleCategories] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[VehicleEntries]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[VehicleEntries] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [VehicleNumber] [varchar](50) NOT NULL,
        [VehicleType] [varchar](50) NULL,
        [DriverName] [varchar](100) NULL,
        [DriverCNIC] [varchar](50) NULL,
        [DriverMobile] [varchar](50) NULL,
        [SupplierCompany] [varchar](150) NULL,
        [Transporter] [varchar](150) NULL,
        [Purpose] [varchar](255) NULL,
        [CategoryId] [int] NULL,
        [EntryTime] [datetime] NOT NULL CONSTRAINT [DF__VehicleEn__Entry__41F0F821] DEFAULT (getdate()),
        [ExitTime] [datetime] NULL,
        [VehiclePhotoBase64] [varchar](MAX) NULL,
        [Status] [varchar](20) NOT NULL CONSTRAINT [DF__VehicleEn__Statu__42E51C5A] DEFAULT ('Inside'),
        [CreatedBy] [varchar](50) NULL,
        [UpdatedBy] [varchar](50) NULL,
        [CreatedAt] [datetime] NOT NULL CONSTRAINT [DF__VehicleEn__Creat__43D94093] DEFAULT (getdate()),
        [UpdatedAt] [datetime] NULL,
        [Remarks] [varchar](255) NULL,
        CONSTRAINT [PK_VehicleEntries] PRIMARY KEY CLUSTERED ([EntryID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[VenderAssItems_Revisions]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[VenderAssItems_Revisions] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [VAI_RefID] [int] NOT NULL,
        [Rate] [float] NULL,
        [Remarks] [varchar](255) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__VenderAss__DTEnt__35361B81] DEFAULT (getdate())
    );
END
GO
IF OBJECT_ID(N'[dbo].[VendRcvdDetail_QC_Report]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[VendRcvdDetail_QC_Report] (
        [VRD_RefID] [int] NOT NULL,
        [TrayNo] [varchar](50) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__VendRcvdD__DTEnt__3FB3A9F4] DEFAULT (getdate()),
        [PassQty] [int] NULL,
        CONSTRAINT [PK_VendRcvdDetail_QC_Report] PRIMARY KEY CLUSTERED ([VRD_RefID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[VendRcvdDetail_QC_Report_PIP]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[VendRcvdDetail_QC_Report_PIP] (
        [VRD_RefID] [int] NOT NULL,
        [PIP_RefID] [int] NOT NULL,
        [DT] [datetime] NULL,
        [EmpID] [varchar](50) NOT NULL,
        [Pieces] [int] NULL,
        [Remarks] [varchar](255) NULL,
        [Action] [tinyint] NULL,
        CONSTRAINT [PK_VendRcvdDetail_QC_Report_PIP] PRIMARY KEY CLUSTERED ([VRD_RefID] ASC, [PIP_RefID] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[VendRcvdDetailPO_Hidden_Billing]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[VendRcvdDetailPO_Hidden_Billing] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [VRDPO_RefID] [int] NOT NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__VendRcvdD__DTEnt__43843AD8] DEFAULT (getdate())
    );
END
GO
IF OBJECT_ID(N'[dbo].[Visitors]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Visitors] (
        [EntryID] [int] IDENTITY(1, 1) NOT NULL,
        [VisitorName] [varchar](50) NULL,
        [CompanyName] [varchar](50) NULL,
        [NICNumber] [varchar](50) NULL,
        [EmpID] [varchar](50) NULL,
        [DT] [datetime] NULL,
        [Purpose] [varchar](255) NULL,
        [Remarks] [varchar](255) NULL,
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [DTEntry] [datetime] NULL CONSTRAINT [DF__Visitors__DTEntr__45C18905] DEFAULT (getdate()),
        [TimeOut] [datetime] NULL,
        [ConcernedPerson] [varchar](50) NULL,
        [MobileNo] [varchar](50) NULL,
        [CardNo] [varchar](50) NULL,
        [VendID] [int] NULL,
        [EmpNameManual] [varchar](50) NULL,
        [VendNameManual] [varchar](50) NULL,
        [ConcernedPersonEmail] [varchar](150) NULL,
        [ConcernedPersonWhatsApp] [varchar](50) NULL,
        [NotifyByEmail] [bit] NOT NULL CONSTRAINT [DF__Visitors__Notify__46B5AD3E] DEFAULT ((0)),
        [NotifyByWhatsApp] [bit] NOT NULL CONSTRAINT [DF__Visitors__Notify__47A9D177] DEFAULT ((0)),
        [VisitCode] [varchar](50) NULL,
        [QrPayload] [varchar](255) NULL,
        [LastNotificationAt] [datetime] NULL,
        [NotificationStatus] [varchar](50) NULL,
        [OutVerifiedBy] [varchar](50) NULL,
        [OutVerifiedAt] [datetime] NULL,
        [OutVerificationInput] [varchar](100) NULL
    );
END
GO
IF OBJECT_ID(N'[dbo].[WorkflowConfigurations]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[WorkflowConfigurations] (
        [WorkflowCode] [varchar](50) NOT NULL,
        [WorkflowName] [varchar](100) NOT NULL,
        [Module] [varchar](50) NOT NULL,
        [Description] [varchar](255) NULL,
        [IsEnabled] [bit] NOT NULL CONSTRAINT [DF_WFConfig_IsEnabled] DEFAULT ((1)),
        [ApproverRole] [varchar](50) NOT NULL CONSTRAINT [DF_WFConfig_ApproverRole] DEFAULT ('Director'),
        [UpdatedDate] [datetime] NOT NULL CONSTRAINT [DF_WFConfig_UpdatedDate] DEFAULT (getdate()),
        [UpdatedBy] [varchar](50) NULL,
        CONSTRAINT [PK_WorkflowConfigurations] PRIMARY KEY CLUSTERED ([WorkflowCode] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[WorkflowExemptRoles]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[WorkflowExemptRoles] (
        [ID] [int] IDENTITY(1, 1) NOT NULL,
        [WorkflowCode] [varchar](50) NOT NULL,
        [RoleName] [varchar](50) NOT NULL,
        CONSTRAINT [PK_WorkflowExemptRoles] PRIMARY KEY CLUSTERED ([ID] ASC),
        CONSTRAINT [UQ_WorkflowExemptRoles_Code_Role] UNIQUE NONCLUSTERED ([WorkflowCode] ASC, [RoleName] ASC)
    );
END
GO
IF OBJECT_ID(N'[dbo].[WorkflowSteps]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[WorkflowSteps] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [WorkflowTemplateId] [int] NOT NULL,
        [StepOrder] [int] NOT NULL,
        [StepName] [nvarchar](MAX) NOT NULL,
        [ApproverType] [int] NOT NULL,
        [SpecificUserId] [nvarchar](100) NULL,
        [RoleName] [nvarchar](MAX) NULL,
        [RequiredAction] [int] NOT NULL,
        [SLAHours] [int] NOT NULL CONSTRAINT [DF__WorkflowS__SLAHo__6702813F] DEFAULT ((24)),
        CONSTRAINT [PK__Workflow__3214EC07C17954CF] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[WorkflowSteps]') AND name = N'IX_WorkflowSteps_WorkflowTemplateId')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_WorkflowSteps_WorkflowTemplateId] ON [dbo].[WorkflowSteps] ([WorkflowTemplateId] ASC);
END
GO
IF OBJECT_ID(N'[dbo].[WorkflowTemplates]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[WorkflowTemplates] (
        [Id] [int] IDENTITY(1, 1) NOT NULL,
        [Name] [nvarchar](400) NOT NULL,
        [Description] [nvarchar](MAX) NULL,
        [RequestType] [nvarchar](200) NOT NULL,
        [IsActive] [bit] NOT NULL CONSTRAINT [DF__WorkflowT__IsAct__6331F05B] DEFAULT ((1)),
        [CreatedAt] [datetime2](7) NOT NULL CONSTRAINT [DF__WorkflowT__Creat__64261494] DEFAULT (getutcdate()),
        [UpdatedAt] [datetime2](7) NULL,
        CONSTRAINT [PK__Workflow__3214EC078BD3B732] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO

-- =============================================================================
-- PHASE 3: MISSING COLUMNS ON EXISTING TABLES
-- =============================================================================
IF COL_LENGTH(N'[dbo].[Accounts]', N'BudgetAccount') IS NULL
BEGIN
    ALTER TABLE [dbo].[Accounts] ADD [BudgetAccount] [bit] NOT NULL CONSTRAINT [DF__Accounts__Budget__5BE5D44E] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Accounts]', N'BudgetAccount_Glance') IS NULL
BEGIN
    ALTER TABLE [dbo].[Accounts] ADD [BudgetAccount_Glance] [bit] NOT NULL CONSTRAINT [DF__Accounts__Budget__5CD9F887] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Accounts]', N'IsNonTransaction') IS NULL
BEGIN
    ALTER TABLE [dbo].[Accounts] ADD [IsNonTransaction] [bit] NULL CONSTRAINT [DF__Accounts__IsNonT__5DCE1CC0] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Accounts]', N'ShowInExpenseReport') IS NULL
BEGIN
    ALTER TABLE [dbo].[Accounts] ADD [ShowInExpenseReport] [bit] NULL CONSTRAINT [DF__Accounts__ShowIn__5EC240F9] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Accounts]', N'ShowInBalanceSheet') IS NULL
BEGIN
    ALTER TABLE [dbo].[Accounts] ADD [ShowInBalanceSheet] [bit] NULL CONSTRAINT [DF__Accounts__ShowIn__5FB66532] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Accounts]', N'SpecialAccount') IS NULL
BEGIN
    ALTER TABLE [dbo].[Accounts] ADD [SpecialAccount] [bit] NULL CONSTRAINT [DF__Accounts__Specia__60AA896B] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[AttendanceMachine]', N'Clientkey') IS NULL
BEGIN
    ALTER TABLE [dbo].[AttendanceMachine] ADD [Clientkey] [varchar](50) NULL CONSTRAINT [DF__Attendanc__Clien__5503CC95] DEFAULT ('Default') WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Company]', N'Employees_Prefix') IS NULL
BEGIN
    ALTER TABLE [dbo].[Company] ADD [Employees_Prefix] [varchar](5) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[CustomerQuotationsItems]', N'ItemRemarks') IS NULL
BEGIN
    ALTER TABLE [dbo].[CustomerQuotationsItems] ADD [ItemRemarks] [varchar](MAX) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[CustomInvoice]', N'GDNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[CustomInvoice] ADD [GDNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[CustomInvoice]', N'InvoiceTo') IS NULL
BEGIN
    ALTER TABLE [dbo].[CustomInvoice] ADD [InvoiceTo] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[DispatchListDetail_VRD]', N'EntryDT') IS NULL
BEGIN
    ALTER TABLE [dbo].[DispatchListDetail_VRD] ADD [EntryDT] [datetime] NULL CONSTRAINT [DF_DispatchListDetail_VRD_EntryDT] DEFAULT (getdate()) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[DispatchListDetail_VRD]', N'AddedBy') IS NULL
BEGIN
    ALTER TABLE [dbo].[DispatchListDetail_VRD] ADD [AddedBy] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[EmployeeAttendanceCatagories]', N'Saturday_Early_Out_Mins') IS NULL
BEGIN
    ALTER TABLE [dbo].[EmployeeAttendanceCatagories] ADD [Saturday_Early_Out_Mins] [int] NOT NULL CONSTRAINT [DF_EmployeeAttendanceCatagories_Saturday_Early_Out_Mins] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Employees]', N'Caste') IS NULL
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [Caste] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Employees]', N'CNIC_PDF') IS NULL
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [CNIC_PDF] [image] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Employees]', N'CNIC_PDF_FileName') IS NULL
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [CNIC_PDF_FileName] [varchar](1000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Employees]', N'Maslak') IS NULL
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [Maslak] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Employees]', N'Pay_Full_Salary') IS NULL
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [Pay_Full_Salary] [bit] NOT NULL CONSTRAINT [DF_Employees_Pay_Full_Salary] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Employees_ST_Sheet_Posting]', N'Posting_Type') IS NULL
BEGIN
    ALTER TABLE [dbo].[Employees_ST_Sheet_Posting] ADD [Posting_Type] [tinyint] NOT NULL CONSTRAINT [DF_Employees_ST_Sheet_Posting_Posting_Type] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FCustomerCatalog]', N'MDMA') IS NULL
BEGIN
    ALTER TABLE [dbo].[FCustomerCatalog] ADD [MDMA] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FCustomerCatalog]', N'SFDA_Listing_No') IS NULL
BEGIN
    ALTER TABLE [dbo].[FCustomerCatalog] ADD [SFDA_Listing_No] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FCustomerCatalog]', N'MD_Group') IS NULL
BEGIN
    ALTER TABLE [dbo].[FCustomerCatalog] ADD [MD_Group] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FCustomerOrders]', N'Packaging_Weight') IS NULL
BEGIN
    ALTER TABLE [dbo].[FCustomerOrders] ADD [Packaging_Weight] [varchar](4000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FCustomerOrders]', N'Authorized') IS NULL
BEGIN
    ALTER TABLE [dbo].[FCustomerOrders] ADD [Authorized] [bit] NOT NULL CONSTRAINT [DF_FCustomerOrders_Authorized] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FCustomerOrders]', N'AuthorizedBy') IS NULL
BEGIN
    ALTER TABLE [dbo].[FCustomerOrders] ADD [AuthorizedBy] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FCustomerOrders]', N'AuthorizedDT') IS NULL
BEGIN
    ALTER TABLE [dbo].[FCustomerOrders] ADD [AuthorizedDT] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'AcumulatedDepriciation') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [AcumulatedDepriciation] [real] NULL CONSTRAINT [DF__FixedAsse__Acumu__6292D1DD] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'DepriciationRate') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [DepriciationRate] [real] NULL CONSTRAINT [DF__FixedAsse__Depri__6386F616] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'VRDPO_RefID') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [VRDPO_RefID] [int] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'OldCode') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [OldCode] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'Disposed') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [Disposed] [bit] NULL CONSTRAINT [DF__FixedAsse__Dispo__647B1A4F] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'DTDisposal') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [DTDisposal] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'SalePrice') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [SalePrice] [real] NULL CONSTRAINT [DF__FixedAsse__SaleP__656F3E88] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'WDV') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [WDV] [real] NULL CONSTRAINT [DF__FixedAssets__WDV__666362C1] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'Remarks') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [Remarks] [varchar](1000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'From_JV') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [From_JV] [bit] NULL CONSTRAINT [DF__FixedAsse__From___675786FA] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'VchrNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [VchrNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'Disposed_Type') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [Disposed_Type] [tinyint] NULL CONSTRAINT [DF__FixedAsse__Dispo__684BAB33] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'EmpID') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [EmpID] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'Approved') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [Approved] [tinyint] NULL CONSTRAINT [DF__FixedAsse__Appro__693FCF6C] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'RcvProcessDT') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [RcvProcessDT] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'UserName') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [UserName] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'MachineName') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [MachineName] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'CaseClosed') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [CaseClosed] [tinyint] NULL CONSTRAINT [DF__FixedAsse__CaseC__6A33F3A5] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'CC_UserName') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [CC_UserName] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'CC_MachineName') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [CC_MachineName] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'CC_EntryDT') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [CC_EntryDT] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'EntryDT') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [EntryDT] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets]', N'InActiveDT') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets] ADD [InActiveDT] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets_Master]', N'PercentRate') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets_Master] ADD [PercentRate] [real] NULL CONSTRAINT [DF__FixedAsse__Perce__6B2817DE] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets_Master]', N'AccumAccNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets_Master] ADD [AccumAccNo] [nvarchar](510) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets_Master]', N'DepriciationAccNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets_Master] ADD [DepriciationAccNo] [nvarchar](510) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets_Master]', N'DepriciationMethodID') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets_Master] ADD [DepriciationMethodID] [tinyint] NULL CONSTRAINT [DF__FixedAsse__Depri__6C1C3C17] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets_Master]', N'Tag') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets_Master] ADD [Tag] [tinyint] NULL CONSTRAINT [DF__FixedAssets__Tag__6D106050] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets_Master]', N'COA_AccNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets_Master] ADD [COA_AccNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets_Master]', N'DepAccNo_Dr') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets_Master] ADD [DepAccNo_Dr] [nvarchar](100) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FixedAssets_Master]', N'DepAccNo_Cr') IS NULL
BEGIN
    ALTER TABLE [dbo].[FixedAssets_Master] ADD [DepAccNo_Cr] [nvarchar](100) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FOrderItems]', N'IW_BatchNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [IW_BatchNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FOrderItems]', N'IW_OrderNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [IW_OrderNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FOrderItems]', N'Authorized') IS NULL
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [Authorized] [bit] NOT NULL CONSTRAINT [DF_FOrderItems_Authorized] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FOrderItems]', N'AuthorizedQty') IS NULL
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [AuthorizedQty] [int] NOT NULL CONSTRAINT [DF_FOrderItems_AuthorizedQty] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[FOrderItems]', N'AuthorizedBy') IS NULL
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [AuthorizedBy] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FOrderItems]', N'AuthorizedDT') IS NULL
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [AuthorizedDT] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[ForeignCustomers]', N'FC_Note_I') IS NULL
BEGIN
    ALTER TABLE [dbo].[ForeignCustomers] ADD [FC_Note_I] [varchar](4000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[ForeignCustomers]', N'FC_Note_II') IS NULL
BEGIN
    ALTER TABLE [dbo].[ForeignCustomers] ADD [FC_Note_II] [varchar](4000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[ForeignCustomers]', N'Inner_Label_Manual_I') IS NULL
BEGIN
    ALTER TABLE [dbo].[ForeignCustomers] ADD [Inner_Label_Manual_I] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[ForeignCustomers]', N'Inner_Label_Manual_II') IS NULL
BEGIN
    ALTER TABLE [dbo].[ForeignCustomers] ADD [Inner_Label_Manual_II] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[FPInvoice]', N'InvoiceTo') IS NULL
BEGIN
    ALTER TABLE [dbo].[FPInvoice] ADD [InvoiceTo] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Heads]', N'BudgetAccount_Glance') IS NULL
BEGIN
    ALTER TABLE [dbo].[Heads] ADD [BudgetAccount_Glance] [bit] NOT NULL CONSTRAINT [DF__Heads__BudgetAcc__619EADA4] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Items]', N'Fig_No') IS NULL
BEGIN
    ALTER TABLE [dbo].[Items] ADD [Fig_No] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Items]', N'Scale') IS NULL
BEGIN
    ALTER TABLE [dbo].[Items] ADD [Scale] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Items]', N'Additional_Detail') IS NULL
BEGIN
    ALTER TABLE [dbo].[Items] ADD [Additional_Detail] [varchar](4000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Items]', N'MakerDescription') IS NULL
BEGIN
    ALTER TABLE [dbo].[Items] ADD [MakerDescription] [varchar](4000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Items]', N'MasterCartonTotalPcs') IS NULL
BEGIN
    ALTER TABLE [dbo].[Items] ADD [MasterCartonTotalPcs] [int] NOT NULL CONSTRAINT [DF_Items_MasterCartonTotalPcs] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Items]', N'TipCap') IS NULL
BEGIN
    ALTER TABLE [dbo].[Items] ADD [TipCap] [varchar](1000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[ItemsRMComp]', N'Functional_Status') IS NULL
BEGIN
    ALTER TABLE [dbo].[ItemsRMComp] ADD [Functional_Status] [tinyint] NOT NULL CONSTRAINT [DF_ItemsRMComp_Functional_Status] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Lots_List]', N'Forge_Batch_No') IS NULL
BEGIN
    ALTER TABLE [dbo].[Lots_List] ADD [Forge_Batch_No] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Lots_List]', N'Lot_Remarks') IS NULL
BEGIN
    ALTER TABLE [dbo].[Lots_List] ADD [Lot_Remarks] [varchar](4000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Makers]', N'CompanyName') IS NULL
BEGIN
    ALTER TABLE [dbo].[Makers] ADD [CompanyName] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Makers]', N'Maker_Second_Name') IS NULL
BEGIN
    ALTER TABLE [dbo].[Makers] ADD [Maker_Second_Name] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Makers]', N'CNIC_PDF') IS NULL
BEGIN
    ALTER TABLE [dbo].[Makers] ADD [CNIC_PDF] [image] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Makers]', N'CNIC_PDF_FileName') IS NULL
BEGIN
    ALTER TABLE [dbo].[Makers] ADD [CNIC_PDF_FileName] [varchar](1000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MaterialMovements]', N'QtyMoved') IS NULL
BEGIN
    ALTER TABLE [dbo].[MaterialMovements] ADD [QtyMoved] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'EmpID') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [EmpID] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'GatePassType') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [GatePassType] [int] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'DeptID') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [DeptID] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'Approved_HOD') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [Approved_HOD] [bit] NULL CONSTRAINT [DF_MiscEntries_GateOut_Approved_HOD] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'Approved') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [Approved] [bit] NULL CONSTRAINT [DF_MiscEntries_GateOut_Approved] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'Approved_Finance') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [Approved_Finance] [bit] NULL CONSTRAINT [DF_MiscEntries_GateOut_Approved_Finance] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'VendID') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [VendID] [int] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'VendNameManual') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [VendNameManual] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut]', N'VM_AccNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut] ADD [VM_AccNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MiscEntries_GateOut_Details]', N'Rate') IS NULL
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut_Details] ADD [Rate] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MonthlySalaries]', N'OTHrs_Original') IS NULL
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [OTHrs_Original] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MonthlySalaries]', N'LateHrs_Original') IS NULL
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [LateHrs_Original] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MonthlySalaries]', N'OTHrs_Net') IS NULL
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [OTHrs_Net] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MonthlySalaries]', N'LateHrs_Net') IS NULL
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [LateHrs_Net] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MonthlySalaries]', N'Salary_Paid') IS NULL
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [Salary_Paid] [bit] NOT NULL CONSTRAINT [DF_MonthlySalaries_Salary_Paid] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[MonthlySalaries]', N'Salary_Paid_UserName') IS NULL
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [Salary_Paid_UserName] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MonthlySalaries]', N'Salary_Paid_MachineName') IS NULL
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [Salary_Paid_MachineName] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[MonthlySalaries]', N'Salary_Paid_DTEntry') IS NULL
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [Salary_Paid_DTEntry] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[PrintSalary]', N'LateHrs_Net') IS NULL
BEGIN
    ALTER TABLE [dbo].[PrintSalary] ADD [LateHrs_Net] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[PrintSalary]', N'LateHrs_Original') IS NULL
BEGIN
    ALTER TABLE [dbo].[PrintSalary] ADD [LateHrs_Original] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[PrintSalary]', N'OTHrs_Net') IS NULL
BEGIN
    ALTER TABLE [dbo].[PrintSalary] ADD [OTHrs_Net] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[PrintSalary]', N'OTHrs_Original') IS NULL
BEGIN
    ALTER TABLE [dbo].[PrintSalary] ADD [OTHrs_Original] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Processes]', N'EstimatedMinutes') IS NULL
BEGIN
    ALTER TABLE [dbo].[Processes] ADD [EstimatedMinutes] [int] NOT NULL CONSTRAINT [DF_Processes_EstimatedMinutes] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[RawMaterialIssuance]', N'Entry_Type') IS NULL
BEGIN
    ALTER TABLE [dbo].[RawMaterialIssuance] ADD [Entry_Type] [tinyint] NOT NULL CONSTRAINT [DF_RawMaterialIssuance_Entry_Type] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[RawMaterialIssuance]', N'Vendor_AccNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[RawMaterialIssuance] ADD [Vendor_AccNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[RawMaterialIssuance]', N'Vendor_PO_No') IS NULL
BEGIN
    ALTER TABLE [dbo].[RawMaterialIssuance] ADD [Vendor_PO_No] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[RawMaterialIssuance]', N'Vendor_VchrNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[RawMaterialIssuance] ADD [Vendor_VchrNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[RawMaterialIssuance]', N'FCO_OrderNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[RawMaterialIssuance] ADD [FCO_OrderNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[RM]', N'RM_Maker_Description') IS NULL
BEGIN
    ALTER TABLE [dbo].[RM] ADD [RM_Maker_Description] [varchar](1000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[RM]', N'Weight_Length') IS NULL
BEGIN
    ALTER TABLE [dbo].[RM] ADD [Weight_Length] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[RunningLots_Temp]', N'TrayNo') IS NULL
BEGIN
    ALTER TABLE [dbo].[RunningLots_Temp] ADD [TrayNo] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'EmpID') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [EmpID] [varchar](50) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'IntraOfficeMainLink') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [IntraOfficeMainLink] [bit] NOT NULL CONSTRAINT [DF_Users_IntraOfficeMainLink] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'MIL_Print_Maker_Issuance_Report_Valuewise') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [MIL_Print_Maker_Issuance_Report_Valuewise] [bit] NOT NULL CONSTRAINT [DF_Users_MIL_Print_Maker_Issuance_Report_Valuewise] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'MIL_Print_Master_PO') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [MIL_Print_Master_PO] [bit] NOT NULL CONSTRAINT [DF_Users_MIL_Print_Master_PO] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'Show_Balance_Vendor_List') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [Show_Balance_Vendor_List] [bit] NOT NULL CONSTRAINT [DF_Users_Show_Balance_Vendor_List] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'Show_Values_Company_Catalog') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [Show_Values_Company_Catalog] [bit] NOT NULL CONSTRAINT [DF_Users_Show_Values_Company_Catalog] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'Show_Customer_Order_No') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [Show_Customer_Order_No] [bit] NOT NULL CONSTRAINT [DF_Users_Show_Customer_Order_No] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'SetupMainLink') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [SetupMainLink] [bit] NOT NULL CONSTRAINT [DF_Users_SetupMainLink] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'BudgetingMainLink') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [BudgetingMainLink] [bit] NULL CONSTRAINT [DF__Users__Budgeting__6E048489] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'GateMainLink') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [GateMainLink] [bit] NULL CONSTRAINT [DF__Users__GateMainL__6EF8A8C2] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AWMGateDashboard') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AWMGateDashboard] [bit] NOT NULL CONSTRAINT [DF__Users__AWMGateDa__6FECCCFB] DEFAULT ((1)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AWMGateIn') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AWMGateIn] [bit] NOT NULL CONSTRAINT [DF__Users__AWMGateIn__70E0F134] DEFAULT ((1)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AWMGatePassTypes') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AWMGatePassTypes] [bit] NOT NULL CONSTRAINT [DF__Users__AWMGatePa__71D5156D] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AWMVisitors') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AWMVisitors] [bit] NOT NULL CONSTRAINT [DF__Users__AWMVisito__72C939A6] DEFAULT ((1)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AWMUserRights') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AWMUserRights] [bit] NOT NULL CONSTRAINT [DF__Users__AWMUserRi__73BD5DDF] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AWMVehicleManagement') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AWMVehicleManagement] [bit] NOT NULL CONSTRAINT [DF__Users__AWMVehicl__74B18218] DEFAULT ((1)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AWMMaterialInward') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AWMMaterialInward] [bit] NOT NULL CONSTRAINT [DF__Users__AWMMateri__75A5A651] DEFAULT ((1)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AWMMaterialOutward') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AWMMaterialOutward] [bit] NOT NULL CONSTRAINT [DF__Users__AWMMateri__7699CA8A] DEFAULT ((1)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'GateLastLoginAt') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [GateLastLoginAt] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'GateLastLogoutAt') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [GateLastLogoutAt] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'GateLastSeenAt') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [GateLastSeenAt] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[Users]', N'AllowOutsideAccess') IS NULL
BEGIN
    ALTER TABLE [dbo].[Users] ADD [AllowOutsideAccess] [bit] NOT NULL CONSTRAINT [DF_Users_AllowOutsideAccess] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[VenderAssItems]', N'Remarks') IS NULL
BEGIN
    ALTER TABLE [dbo].[VenderAssItems] ADD [Remarks] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssdDetail]', N'CountedBy') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail] ADD [CountedBy] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssdDetail_ReturnDTs_Revisions]', N'Qty_New') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail_ReturnDTs_Revisions] ADD [Qty_New] [int] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssdDetail_ReturnDTs_Revisions]', N'Qty_Old') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail_ReturnDTs_Revisions] ADD [Qty_Old] [int] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssdDetail_ReturnDTs_Revisions]', N'Rate_New') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail_ReturnDTs_Revisions] ADD [Rate_New] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssdDetail_ReturnDTs_Revisions]', N'Rate_Old') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail_ReturnDTs_Revisions] ADD [Rate_Old] [real] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssued]', N'DTEntry') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssued] ADD [DTEntry] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssued]', N'MakerSignedCopyPath') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssued] ADD [MakerSignedCopyPath] [nvarchar](500) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssued]', N'MakerSignedCopyFileName') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssued] ADD [MakerSignedCopyFileName] [nvarchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssued]', N'MakerSignedCopyUploadedAt') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssued] ADD [MakerSignedCopyUploadedAt] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendIssued]', N'MakerSignedCopyUploadedBy') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendIssued] ADD [MakerSignedCopyUploadedBy] [nvarchar](100) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendOrderDetail]', N'ItemRemarks') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendOrderDetail] ADD [ItemRemarks] [varchar](4000) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VendOrders]', N'Order_Collection_By') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendOrders] ADD [Order_Collection_By] [varchar](255) NULL;
END
GO
IF COL_LENGTH(N'[dbo].[vendRcvdDetail]', N'Not_Available_For_Billing') IS NULL
BEGIN
    ALTER TABLE [dbo].[vendRcvdDetail] ADD [Not_Available_For_Billing] [bit] NOT NULL CONSTRAINT [DF_vendRcvdDetail_Not_Available_For_Billing] DEFAULT ((0)) WITH VALUES;
END
GO
IF COL_LENGTH(N'[dbo].[VendReceived]', N'EntryDT') IS NULL
BEGIN
    ALTER TABLE [dbo].[VendReceived] ADD [EntryDT] [datetime] NULL;
END
GO
IF COL_LENGTH(N'[dbo].[VoucherInfo]', N'Handed_Over_To') IS NULL
BEGIN
    ALTER TABLE [dbo].[VoucherInfo] ADD [Handed_Over_To] [varchar](255) NULL;
END
GO

-- =============================================================================
-- PHASE 4: COLUMN TYPE ALTERATIONS ON EXISTING TABLES
-- =============================================================================
IF EXISTS (
    SELECT 1
    FROM sys.columns c
    JOIN sys.types ty ON c.user_type_id = ty.user_type_id
    WHERE c.object_id = OBJECT_ID(N'[dbo].[BalanceTags]') AND c.name = N'Sno' AND ty.name <> N'float'
)
BEGIN
    IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE parent_object_id = OBJECT_ID(N'[dbo].[BalanceTags]') AND name = N'PK_BalanceTags')
    BEGIN
        ALTER TABLE [dbo].[BalanceTags] DROP CONSTRAINT [PK_BalanceTags];
    END
    ALTER TABLE [dbo].[BalanceTags] ALTER COLUMN [Sno] [float] NOT NULL;
    IF NOT EXISTS (SELECT 1 FROM sys.key_constraints WHERE parent_object_id = OBJECT_ID(N'[dbo].[BalanceTags]') AND type = 'PK')
    BEGIN
        ALTER TABLE [dbo].[BalanceTags] ADD CONSTRAINT [PK_BalanceTags] PRIMARY KEY CLUSTERED ([Sno] ASC);
    END
END
GO
IF EXISTS (
    SELECT 1
    FROM sys.columns c
    JOIN sys.types ty ON c.user_type_id = ty.user_type_id
    WHERE c.object_id = OBJECT_ID(N'[dbo].[MiscEntries_GateOut_Details]') AND c.name = N'Qty' AND (ty.name <> N'real' OR c.is_nullable = 0)
)
BEGIN
    ALTER TABLE [dbo].[MiscEntries_GateOut_Details] ALTER COLUMN [Qty] [real] NULL;
END
GO

-- =============================================================================
-- PHASE 5: MISSING DEFAULT CONSTRAINTS, INDEXES & FOREIGN KEYS
-- =============================================================================
IF COL_LENGTH(N'[dbo].[VendAssItems]', N'CustCode') IS NOT NULL
   AND NOT EXISTS (
       SELECT 1 FROM sys.default_constraints
       WHERE parent_object_id = OBJECT_ID(N'[dbo].[VendAssItems]')
         AND parent_column_id = COLUMNPROPERTY(OBJECT_ID(N'[dbo].[VendAssItems]'), N'CustCode', 'ColumnId')
   )
BEGIN
    ALTER TABLE [dbo].[VendAssItems] ADD CONSTRAINT [DF_VendAssItems_CustCode] DEFAULT ('Stock') FOR [CustCode];
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[vendRcvdDetail]') AND name = N'IX_vendRcvdDetail_RefID_ItemCode')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_vendRcvdDetail_RefID_ItemCode] ON [dbo].[vendRcvdDetail] ([RefID] ASC, [ItemCode] ASC) INCLUDE ([EntryID], [RecieptID], [LotNo], [OrderNo], [RcvdQty], [Wastage], [LostQty], [ReWorkQty], [ReWorkLot], [IssQty], [Opening_RefID], [Issue_RefID]);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[VendReceived]') AND name = N'IX_VendReceived_DT')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_VendReceived_DT] ON [dbo].[VendReceived] ([DT] ASC) INCLUDE ([EntryID], [VendID], [ProcessID], [Issuance_RefID], [EmpID], [UserID]);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[VendReceived_Employees]') AND name = N'IX_VendReceived_Employees_VR_RefID')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_VendReceived_Employees_VR_RefID] ON [dbo].[VendReceived_Employees] ([VR_RefID] ASC, [EmpID] ASC);
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Vouchers]') AND name = N'IX_Vouchers_AccNo_SNo')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Vouchers_AccNo_SNo] ON [dbo].[Vouchers] ([Accno] ASC, [VDate] ASC) INCLUDE ([Debit], [Credit], [balance]);
END
GO
IF OBJECT_ID(N'[dbo].[AnnouncementAcknowledgments]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Announcements]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_AnnouncementAcknowledgments_Announcements' AND parent_object_id = OBJECT_ID(N'[dbo].[AnnouncementAcknowledgments]'))
BEGIN
    ALTER TABLE [dbo].[AnnouncementAcknowledgments] WITH CHECK ADD CONSTRAINT [FK_AnnouncementAcknowledgments_Announcements] FOREIGN KEY ([AnnouncementId]) REFERENCES [dbo].[Announcements] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[AnnouncementAcknowledgments] CHECK CONSTRAINT [FK_AnnouncementAcknowledgments_Announcements];
END
GO
IF OBJECT_ID(N'[dbo].[AnnouncementAttachments]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Announcements]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_AnnouncementAttachments_Announcements' AND parent_object_id = OBJECT_ID(N'[dbo].[AnnouncementAttachments]'))
BEGIN
    ALTER TABLE [dbo].[AnnouncementAttachments] WITH CHECK ADD CONSTRAINT [FK_AnnouncementAttachments_Announcements] FOREIGN KEY ([AnnouncementId]) REFERENCES [dbo].[Announcements] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[AnnouncementAttachments] CHECK CONSTRAINT [FK_AnnouncementAttachments_Announcements];
END
GO
IF OBJECT_ID(N'[dbo].[ApprovalRequestAttachments]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[ApprovalRequests]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ApprovalRequestAttachments_ApprovalRequests' AND parent_object_id = OBJECT_ID(N'[dbo].[ApprovalRequestAttachments]'))
BEGIN
    ALTER TABLE [dbo].[ApprovalRequestAttachments] WITH CHECK ADD CONSTRAINT [FK_ApprovalRequestAttachments_ApprovalRequests] FOREIGN KEY ([ApprovalRequestId]) REFERENCES [dbo].[ApprovalRequests] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[ApprovalRequestAttachments] CHECK CONSTRAINT [FK_ApprovalRequestAttachments_ApprovalRequests];
END
GO
IF OBJECT_ID(N'[dbo].[ApprovalRequestHistories]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[ApprovalRequests]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ApprovalRequestHistories_ApprovalRequests' AND parent_object_id = OBJECT_ID(N'[dbo].[ApprovalRequestHistories]'))
BEGIN
    ALTER TABLE [dbo].[ApprovalRequestHistories] WITH CHECK ADD CONSTRAINT [FK_ApprovalRequestHistories_ApprovalRequests] FOREIGN KEY ([ApprovalRequestId]) REFERENCES [dbo].[ApprovalRequests] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[ApprovalRequestHistories] CHECK CONSTRAINT [FK_ApprovalRequestHistories_ApprovalRequests];
END
GO
IF OBJECT_ID(N'[dbo].[ApprovalRequests]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[WorkflowTemplates]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ApprovalRequests_WorkflowTemplates' AND parent_object_id = OBJECT_ID(N'[dbo].[ApprovalRequests]'))
BEGIN
    ALTER TABLE [dbo].[ApprovalRequests] WITH CHECK ADD CONSTRAINT [FK_ApprovalRequests_WorkflowTemplates] FOREIGN KEY ([WorkflowTemplateId]) REFERENCES [dbo].[WorkflowTemplates] ([Id]);
    ALTER TABLE [dbo].[ApprovalRequests] CHECK CONSTRAINT [FK_ApprovalRequests_WorkflowTemplates];
END
GO
IF OBJECT_ID(N'[dbo].[ChannelMembers]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Channels]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ChannelMembers_Channels' AND parent_object_id = OBJECT_ID(N'[dbo].[ChannelMembers]'))
BEGIN
    ALTER TABLE [dbo].[ChannelMembers] WITH CHECK ADD CONSTRAINT [FK_ChannelMembers_Channels] FOREIGN KEY ([ChannelId]) REFERENCES [dbo].[Channels] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[ChannelMembers] CHECK CONSTRAINT [FK_ChannelMembers_Channels];
END
GO
IF OBJECT_ID(N'[dbo].[Items_Parts]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Items]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Items_Parts_Items' AND parent_object_id = OBJECT_ID(N'[dbo].[Items_Parts]'))
BEGIN
    ALTER TABLE [dbo].[Items_Parts] WITH NOCHECK ADD CONSTRAINT [FK_Items_Parts_Items] FOREIGN KEY ([ItemID]) REFERENCES [dbo].[Items] ([ItemID]) ON DELETE CASCADE ON UPDATE CASCADE;
END
GO
IF OBJECT_ID(N'[dbo].[MeetingParticipants]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Meetings]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_MeetingParticipants_Meetings' AND parent_object_id = OBJECT_ID(N'[dbo].[MeetingParticipants]'))
BEGIN
    ALTER TABLE [dbo].[MeetingParticipants] WITH CHECK ADD CONSTRAINT [FK_MeetingParticipants_Meetings] FOREIGN KEY ([MeetingId]) REFERENCES [dbo].[Meetings] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[MeetingParticipants] CHECK CONSTRAINT [FK_MeetingParticipants_Meetings];
END
GO
IF OBJECT_ID(N'[dbo].[MessageAttachments]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Messages]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_MessageAttachments_Messages' AND parent_object_id = OBJECT_ID(N'[dbo].[MessageAttachments]'))
BEGIN
    ALTER TABLE [dbo].[MessageAttachments] WITH CHECK ADD CONSTRAINT [FK_MessageAttachments_Messages] FOREIGN KEY ([MessageId]) REFERENCES [dbo].[Messages] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[MessageAttachments] CHECK CONSTRAINT [FK_MessageAttachments_Messages];
END
GO
IF OBJECT_ID(N'[dbo].[Messages]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Channels]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Messages_Channels' AND parent_object_id = OBJECT_ID(N'[dbo].[Messages]'))
BEGIN
    ALTER TABLE [dbo].[Messages] WITH CHECK ADD CONSTRAINT [FK_Messages_Channels] FOREIGN KEY ([ChannelId]) REFERENCES [dbo].[Channels] ([Id]) ON DELETE SET NULL;
    ALTER TABLE [dbo].[Messages] CHECK CONSTRAINT [FK_Messages_Channels];
END
GO
IF OBJECT_ID(N'[dbo].[Messages]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Messages]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Messages_ParentMessage' AND parent_object_id = OBJECT_ID(N'[dbo].[Messages]'))
BEGIN
    ALTER TABLE [dbo].[Messages] WITH CHECK ADD CONSTRAINT [FK_Messages_ParentMessage] FOREIGN KEY ([ParentMessageId]) REFERENCES [dbo].[Messages] ([Id]);
    ALTER TABLE [dbo].[Messages] CHECK CONSTRAINT [FK_Messages_ParentMessage];
END
GO
IF OBJECT_ID(N'[dbo].[MinuteAttachments]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[MinuteApprovals]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_MinuteAttachments_MinuteApprovals' AND parent_object_id = OBJECT_ID(N'[dbo].[MinuteAttachments]'))
BEGIN
    ALTER TABLE [dbo].[MinuteAttachments] WITH CHECK ADD CONSTRAINT [FK_MinuteAttachments_MinuteApprovals] FOREIGN KEY ([MinuteApprovalId]) REFERENCES [dbo].[MinuteApprovals] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[MinuteAttachments] CHECK CONSTRAINT [FK_MinuteAttachments_MinuteApprovals];
END
GO
IF OBJECT_ID(N'[dbo].[MinuteWorkflowHistories]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[MinuteApprovals]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_MinuteWorkflowHistories_MinuteApprovals' AND parent_object_id = OBJECT_ID(N'[dbo].[MinuteWorkflowHistories]'))
BEGIN
    ALTER TABLE [dbo].[MinuteWorkflowHistories] WITH CHECK ADD CONSTRAINT [FK_MinuteWorkflowHistories_MinuteApprovals] FOREIGN KEY ([MinuteApprovalId]) REFERENCES [dbo].[MinuteApprovals] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[MinuteWorkflowHistories] CHECK CONSTRAINT [FK_MinuteWorkflowHistories_MinuteApprovals];
END
GO
IF OBJECT_ID(N'[dbo].[PPC_Order_Item_Planning]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[FCustomerOrders]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_PpcItemPlanning_OrderNo' AND parent_object_id = OBJECT_ID(N'[dbo].[PPC_Order_Item_Planning]'))
BEGIN
    ALTER TABLE [dbo].[PPC_Order_Item_Planning] WITH CHECK ADD CONSTRAINT [FK_PpcItemPlanning_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[PPC_Order_Item_Planning] CHECK CONSTRAINT [FK_PpcItemPlanning_OrderNo];
END
GO
IF OBJECT_ID(N'[dbo].[ProcessGroup_Hub_Supervisors]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Hub_Names]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_PG_HubSupervisors_HubNames' AND parent_object_id = OBJECT_ID(N'[dbo].[ProcessGroup_Hub_Supervisors]'))
BEGIN
    ALTER TABLE [dbo].[ProcessGroup_Hub_Supervisors] WITH CHECK ADD CONSTRAINT [FK_PG_HubSupervisors_HubNames] FOREIGN KEY ([Hub_Name]) REFERENCES [dbo].[Hub_Names] ([Hub_Name]) ON UPDATE CASCADE;
    ALTER TABLE [dbo].[ProcessGroup_Hub_Supervisors] CHECK CONSTRAINT [FK_PG_HubSupervisors_HubNames];
END
GO
IF OBJECT_ID(N'[dbo].[ProcessGroup_Hub_Supervisors]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[ProcessGroups]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_PG_HubSupervisors_ProcessGroups' AND parent_object_id = OBJECT_ID(N'[dbo].[ProcessGroup_Hub_Supervisors]'))
BEGIN
    ALTER TABLE [dbo].[ProcessGroup_Hub_Supervisors] WITH CHECK ADD CONSTRAINT [FK_PG_HubSupervisors_ProcessGroups] FOREIGN KEY ([GroupID]) REFERENCES [dbo].[ProcessGroups] ([EntryID]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[ProcessGroup_Hub_Supervisors] CHECK CONSTRAINT [FK_PG_HubSupervisors_ProcessGroups];
END
GO
IF OBJECT_ID(N'[dbo].[ProcessGroup_Hub_Supervisors]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Users]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_PG_HubSupervisors_Users' AND parent_object_id = OBJECT_ID(N'[dbo].[ProcessGroup_Hub_Supervisors]'))
BEGIN
    ALTER TABLE [dbo].[ProcessGroup_Hub_Supervisors] WITH CHECK ADD CONSTRAINT [FK_PG_HubSupervisors_Users] FOREIGN KEY ([UserID]) REFERENCES [dbo].[Users] ([UserID]);
    ALTER TABLE [dbo].[ProcessGroup_Hub_Supervisors] CHECK CONSTRAINT [FK_PG_HubSupervisors_Users];
END
GO
IF OBJECT_ID(N'[dbo].[Task_Assignees]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[TaskItems]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Task_Assignees_TaskItems' AND parent_object_id = OBJECT_ID(N'[dbo].[Task_Assignees]'))
BEGIN
    ALTER TABLE [dbo].[Task_Assignees] WITH CHECK ADD CONSTRAINT [FK_Task_Assignees_TaskItems] FOREIGN KEY ([TaskID]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[Task_Assignees] CHECK CONSTRAINT [FK_Task_Assignees_TaskItems];
END
GO
IF OBJECT_ID(N'[dbo].[Task_Assignees]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Users]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Task_Assignees_Users' AND parent_object_id = OBJECT_ID(N'[dbo].[Task_Assignees]'))
BEGIN
    ALTER TABLE [dbo].[Task_Assignees] WITH CHECK ADD CONSTRAINT [FK_Task_Assignees_Users] FOREIGN KEY ([UserID]) REFERENCES [dbo].[Users] ([UserID]);
    ALTER TABLE [dbo].[Task_Assignees] CHECK CONSTRAINT [FK_Task_Assignees_Users];
END
GO
IF OBJECT_ID(N'[dbo].[Task_Roles]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[TaskItems]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Task_Roles_TaskItems' AND parent_object_id = OBJECT_ID(N'[dbo].[Task_Roles]'))
BEGIN
    ALTER TABLE [dbo].[Task_Roles] WITH CHECK ADD CONSTRAINT [FK_Task_Roles_TaskItems] FOREIGN KEY ([TaskID]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[Task_Roles] CHECK CONSTRAINT [FK_Task_Roles_TaskItems];
END
GO
IF OBJECT_ID(N'[dbo].[TaskAttachments]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[TaskItems]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_TaskAttachments_TaskItems' AND parent_object_id = OBJECT_ID(N'[dbo].[TaskAttachments]'))
BEGIN
    ALTER TABLE [dbo].[TaskAttachments] WITH CHECK ADD CONSTRAINT [FK_TaskAttachments_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[TaskAttachments] CHECK CONSTRAINT [FK_TaskAttachments_TaskItems];
END
GO
IF OBJECT_ID(N'[dbo].[TaskComments]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[TaskItems]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_TaskComments_TaskItems' AND parent_object_id = OBJECT_ID(N'[dbo].[TaskComments]'))
BEGIN
    ALTER TABLE [dbo].[TaskComments] WITH CHECK ADD CONSTRAINT [FK_TaskComments_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[TaskComments] CHECK CONSTRAINT [FK_TaskComments_TaskItems];
END
GO
IF OBJECT_ID(N'[dbo].[Users_User_Roles]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[User_Roles]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Users_User_Roles_User_Roles' AND parent_object_id = OBJECT_ID(N'[dbo].[Users_User_Roles]'))
BEGIN
    ALTER TABLE [dbo].[Users_User_Roles] WITH CHECK ADD CONSTRAINT [FK_Users_User_Roles_User_Roles] FOREIGN KEY ([User_Role]) REFERENCES [dbo].[User_Roles] ([User_Role]) ON DELETE CASCADE ON UPDATE CASCADE;
    ALTER TABLE [dbo].[Users_User_Roles] CHECK CONSTRAINT [FK_Users_User_Roles_User_Roles];
END
GO
IF OBJECT_ID(N'[dbo].[Users_User_Roles]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[Users]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Users_User_Roles_Users' AND parent_object_id = OBJECT_ID(N'[dbo].[Users_User_Roles]'))
BEGIN
    ALTER TABLE [dbo].[Users_User_Roles] WITH CHECK ADD CONSTRAINT [FK_Users_User_Roles_Users] FOREIGN KEY ([UserID]) REFERENCES [dbo].[Users] ([UserID]) ON DELETE CASCADE ON UPDATE CASCADE;
    ALTER TABLE [dbo].[Users_User_Roles] CHECK CONSTRAINT [FK_Users_User_Roles_Users];
END
GO
IF OBJECT_ID(N'[dbo].[WorkflowExemptRoles]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[WorkflowConfigurations]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_WorkflowExemptRoles_Config' AND parent_object_id = OBJECT_ID(N'[dbo].[WorkflowExemptRoles]'))
BEGIN
    ALTER TABLE [dbo].[WorkflowExemptRoles] WITH CHECK ADD CONSTRAINT [FK_WorkflowExemptRoles_Config] FOREIGN KEY ([WorkflowCode]) REFERENCES [dbo].[WorkflowConfigurations] ([WorkflowCode]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[WorkflowExemptRoles] CHECK CONSTRAINT [FK_WorkflowExemptRoles_Config];
END
GO
IF OBJECT_ID(N'[dbo].[WorkflowSteps]', N'U') IS NOT NULL
   AND OBJECT_ID(N'[dbo].[WorkflowTemplates]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_WorkflowSteps_WorkflowTemplates' AND parent_object_id = OBJECT_ID(N'[dbo].[WorkflowSteps]'))
BEGIN
    ALTER TABLE [dbo].[WorkflowSteps] WITH CHECK ADD CONSTRAINT [FK_WorkflowSteps_WorkflowTemplates] FOREIGN KEY ([WorkflowTemplateId]) REFERENCES [dbo].[WorkflowTemplates] ([Id]) ON DELETE CASCADE;
    ALTER TABLE [dbo].[WorkflowSteps] CHECK CONSTRAINT [FK_WorkflowSteps_WorkflowTemplates];
END
GO

-- =============================================================================
-- PHASE 6: FUNCTIONS, VIEWS & STORED PROCEDURES
-- =============================================================================
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn_IsLotPendingDeletion]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[fn_IsLotPendingDeletion];
END
GO

CREATE   FUNCTION [dbo].[fn_IsLotPendingDeletion] (@LotNo VARCHAR(50), @VRD_EntryID BIGINT = 0)
RETURNS BIT
AS
BEGIN
    IF EXISTS (
        SELECT 1 FROM dbo.ProductionDeletionRequests WITH (NOLOCK)
        WHERE Status = 'Pending'
          AND (
              (@VRD_EntryID > 0 AND EntityRefID = @VRD_EntryID)
              OR (@VRD_EntryID = 0 AND LotNo = @LotNo)
          )
    )
        RETURN 1;
    RETURN 0;
END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FSundayOTHrs_BZeeshan]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[FSundayOTHrs_BZeeshan];
END
GO

-- ====================================================================================================
-- PART 4: USER-DEFINED FUNCTIONS
-- ====================================================================================================
CREATE   FUNCTION [dbo].[FSundayOTHrs_BZeeshan](@ParDate datetime)
RETURNS TABLE AS  
RETURN(
SELECT EmpID--, SUM(CASE WHEN PayableHrs>8 THEN 8 ELSE ROUND(PayableHrs,1) END) AS SundayOTHrs
	, SUM(ROUND(PayableHrs,2)+ROUND(OTHrs,2)) AS SundayOTHrs
FROM VEmpTimes2
WHERE (MONTH(dt) = MONTH(@ParDate) AND YEAR(DT) = YEAR(@ParDate) AND DT <=@ParDATE )  AND DATENAME(WEEKDAY,DT)='Sunday'
GROUP BY EmpID)

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Get_FixedAssets_Adjustment]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[Get_FixedAssets_Adjustment];
END
GO


-- ------------------------------------------------------------------------------
-- 4. STORED PROCEDURES
-- ------------------------------------------------------------------------------

CREATE   FUNCTION [dbo].[Get_FixedAssets_Adjustment] (@AssetID AS INT,@DT AS DATETIME)  
	RETURNS REAL AS
	BEGIN
		DECLARE @ReturnVal REAL
		DECLARE @Rate AS REAL

		SELECT @Rate=SUM(ISNULL(Adjustment,0)) FROM FixedAssets_Adjustment WHERE DT<=@DT AND AssetID=@AssetID GROUP BY AssetID
		SET @ReturnVal=@Rate

		RETURN @ReturnVal
	END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GetItem_Previous_Next_ProcessID_TblFn]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[GetItem_Previous_Next_ProcessID_TblFn];
END
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE   FUNCTION GetItem_Previous_Next_ProcessID_TblFn(@ProcessID INT)

RETURNS TABLE 
AS
RETURN 
(
	SELECT T1.* FROM (
		SELECT ItemID
			,LAG(ProcessID) OVER(PARTITION BY ItemID ORDER BY SNo) AS Previous_ProcessID
			,ProcessID AS Current_ProcessID 
			,LEAD(ProcessID) OVER(PARTITION BY ItemID ORDER BY SNo) AS Next_ProcessID
		FROM ItemProcesses --WHERE ProcessID=@ProcessID
		) T1 WHERE Current_ProcessID=@ProcessID
)

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GetNextAccno]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[GetNextAccno];
END
GO

-- 7. Ensure GetNextAccno function handles alphabetical and numeric account prefixes safely
CREATE   FUNCTION [dbo].[GetNextAccno](@AccountName AS VARCHAR(255),@AccType AS VARCHAR(50),@ParentAccount As Varchar(255)=NULL,@IsParent As BIT=0)
RETURNS VARCHAR(255) AS  
BEGIN
DECLARE @AccNo AS VARCHAR(255)
DECLARE @NewVal AS VARCHAR(255), @Prefix As VARCHAR(50)

IF @IsParent=1
	BEGIN
		SET @NewVal=(SELECT MAX(CAST(RIGHT(AccNo,3) AS INT)) from Accounts WHERE SubAccOf=@ParentAccount)
		SET @NewVal=CAST(ISNULL(@NewVal,'') AS INT)+1
		SET @AccNo=@ParentAccount + '-' + REPLICATE('0',3-LEN(@NewVal)) + @NewVal
	END
ELSE
	BEGIN
		SET @Prefix=ASCII(UPPER(LEFT(LTRIM(RTRIM(ISNULL(@AccountName,''))),1)))
		IF @Prefix>=65 AND @Prefix<=90
			SET @Prefix=REPLICATE('0',2-LEN(@Prefix-64)) + CAST(@Prefix-64 AS VARCHAR)
		ELSE 
			SET @Prefix='00'
		
		SET @NewVal=(SELECT MAX(CAST(RIGHT(AccNo,3) AS INT)) FROM Accounts WHERE SubAccOf=@ParentAccount AND SUBSTRING(AccNo,LEN(Accno)-4,2)=@Prefix)
		SET @NewVal=CAST(ISNULL(@NewVal,'') AS INT)+1
		SET @AccNo=@ParentAccount + '-' + @Prefix + REPLICATE('0',3-LEN(@NewVal)) + @NewVal
	END
RETURN @AccNo
END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GetNextVchrNo]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[GetNextVchrNo];
END
GO

CREATE   FUNCTION [dbo].[GetNextVchrNo](@DT DATETIME,@VchrType VARCHAR(50))

RETURNS VARCHAR(255) AS  
	BEGIN
		DECLARE @NewVchrNo INT
		SELECT @NewVchrNo=MAX(CAST(RIGHT(VchrNo,CHARINDEX('-',REVERSE(VchrNo))-1) AS INT)) FROM Vouchers 
			WHERE VDate=@DT
			AND LEFT(VchrNo,8+LEN(@VchrType))=@VchrType+'-'+FORMAT(@DT,'yyMMdd')+'-'

		SET @NewVchrNo=ISNULL(@NewVchrNo,0)+1

	
	RETURN @VchrType+'-'+FORMAT(@DT,'yyMMdd')+'-'+FORMAT(@NewVchrNo,'000')
END


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GetTransactionValue]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[GetTransactionValue];
END
GO

-- ------------------------------------------------------------------------------
-- 2. SCALAR FUNCTIONS
-- ------------------------------------------------------------------------------

CREATE   FUNCTION [dbo].[GetTransactionValue] (@AccNo AS VARCHAR(50),@DTFrom DATETIME,@DTTo DATETIME)  
RETURNS INT AS
BEGIN
    DECLARE @Balance INT;
    SELECT @Balance = SUM(Debit-Credit) 
    FROM Vouchers 
    WHERE VDate BETWEEN @DTFrom AND @DTTo
      AND LEFT(AccNo, LEN(@AccNo)) = @AccNo;
    RETURN ISNULL(@Balance, 0);
END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Invoice_DispatchListNos_Fn]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[Invoice_DispatchListNos_Fn];
END
GO

CREATE   FUNCTION [dbo].[Invoice_DispatchListNos_Fn] (@InvoiceNo VARCHAR(50))
RETURNS VARCHAR(1000) AS

BEGIN
DECLARE @ReturnStr AS VARCHAR(1000)
SET @ReturnStr=''
DECLARE @Temp AS VARCHAR(255)
DECLARE Abc CURSOR FOR SELECT DISTINCT DispatchListNo FROM CustomPList_DispatchListDetail
						INNER JOIN CustomPList ON CustomPList_DispatchListDetail.CustomPList_RefID=CustomPList.ID
						INNER JOIN DispatchList ON CustomPList_DispatchListDetail.DP_RefID=DispatchList.EntryID
						WHERE CustomPList.CustomInvoice=@InvoiceNo
OPEN Abc

FETCH NEXT FROM Abc into @Temp
WHILE @@Fetch_Status=0
	BEGIN
		SET @ReturnStr=@ReturnStr + @Temp + ',' 
		FETCH NEXT FROM Abc Into @Temp
	END
	IF LEN(@ReturnStr)>0 
		SET @ReturnStr = LEFT(@ReturnStr,LEN(@ReturnStr)-1)
CLOSE Abc
DEALLOCATE Abc

RETURN @ReturnStr

END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Lot_State_For_Inner_Labels_F]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[Lot_State_For_Inner_Labels_F];
END
GO

CREATE   FUNCTION [dbo].[Lot_State_For_Inner_Labels_F] (@EntryID AS INT,@ShowAll AS BIT=0)
RETURNS VARCHAR(1000) AS

BEGIN
DECLARE @Desc AS VARCHAR(4000)
SET @Desc=''
DECLARE  @Process AS VARCHAR(255)
/*DECLARE @Temper_Process_ID VARCHAR(50),@Inspection_Process_ID VARCHAR(50)
SET @Temper_Process_ID=dbo.GetGeneralDataValue_F('Temper_Inspection_ProcessID')
SET @Inspection_Process_ID=dbo.GetGeneralDataValue_F('First_Inspection_ProcessID')*/

IF @ShowAll=0

	DECLARE myCursor CURSOR FOR SELECT TOP 1 Processes.Description FROM DispatchListDetail_VRD 
	INNER JOIN VendRcvdDetail ON DispatchListDetail_VRD.VRD_RefID=VendRcvdDetail.EntryID 
	INNER JOIN DispatchListDetails_Adv ON DispatchListDetail_VRD.DLD_refID=DispatchListDetails_Adv.EntryID
	INNER JOIN DispatchList ON DispatchList.EntryID=DispatchListDetails_Adv.RefID
	INNER JOIN DispatchListDetail_Inners ON DispatchListDetails_Adv.EntryID=DispatchListDetail_Inners.RefID	
	INNER JOIN Processes ON VendRcvdDetail.ProcessID=Processes.ProcessID
	WHERE DLDC_RefID=@EntryID AND DispatchList.NewFormat=1 ORDER BY DispatchListDetail_Inners.Qty DESC

ELSE

	DECLARE myCursor CURSOR FOR SELECT Processes.Description FROM DispatchListDetail_VRD 
	INNER JOIN VendRcvdDetail ON DispatchListDetail_VRD.VRD_RefID=VendRcvdDetail.EntryID 
	INNER JOIN DispatchListDetails_Adv ON DispatchListDetail_VRD.DLD_refID=DispatchListDetails_Adv.EntryID
	INNER JOIN DispatchList ON DispatchList.EntryID=DispatchListDetails_Adv.RefID
	INNER JOIN DispatchListDetail_Inners ON DispatchListDetails_Adv.EntryID=DispatchListDetail_Inners.RefID	
	INNER JOIN Processes ON VendRcvdDetail.ProcessID=Processes.ProcessID
	WHERE DLDC_RefID=@EntryID AND DispatchList.NewFormat=1 ORDER BY DispatchListDetail_Inners.Qty DESC

	
OPEN myCursor

FETCH NEXT FROM myCursor INTO @Process
WHILE @@Fetch_Status=0
	BEGIN
		SET @Desc=@Desc +  @Process + ','
		FETCH NEXT FROM myCursor INTO @Process
	END
	IF LEN(@Desc)>0
		SET @Desc = LEFT(@Desc,LEN(@Desc)-1)
CLOSE myCursor
DEALLOCATE myCursor

RETURN @Desc

END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[RunningLots]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[RunningLots];
END
GO

CREATE   FUNCTION [dbo].[RunningLots] (@OrderNo AS VARCHAR (255),@ItemID AS VARCHAR (255))
RETURNS VARCHAR(1000) AS

BEGIN
DECLARE @ReturnStr AS VARCHAR(1000)
SET @ReturnStr=''
DECLARE @LotNo AS VARCHAR(255),@Description AS VARCHAR(255),@Qty INT
--DECLARE Abc CURSOR FOR SELECT LotNo,Description,Qty FROM VRunningLots_Simple WHERE OrderNo=@OrderNo AND ItemCode=@ItemID
DECLARE Abc CURSOR FOR SELECT DISTINCT Description FROM VRunningLots_Simple WHERE OrderNo=@OrderNo AND ItemCode=@ItemID
OPEN Abc

--FETCH NEXT FROM Abc INTO @LotNo,@Description,@Qty
FETCH NEXT FROM Abc INTO @Description
WHILE @@Fetch_Status=0
	BEGIN
		SET @ReturnStr=@ReturnStr + @Description + ',' 
		--FETCH NEXT FROM Abc INTO @LotNo,@Description,@Qty
		FETCH NEXT FROM Abc INTO @Description
	END
	IF LEN(@ReturnStr)>0 
		SET @ReturnStr = LEFT(@ReturnStr,LEN(@ReturnStr)-1)
CLOSE Abc
DEALLOCATE Abc

RETURN @ReturnStr

END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VendReceived_Employees_F]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[VendReceived_Employees_F];
END
GO

CREATE   FUNCTION [dbo].[VendReceived_Employees_F] (@EntryID INT)
RETURNS VARCHAR(1000)
AS
BEGIN
    DECLARE @ReturnStr AS VARCHAR(1000);

    SELECT @ReturnStr = STUFF((
        SELECT ', ' + E.Name
        FROM dbo.VendReceived_Employees VRE
        INNER JOIN dbo.Employees E ON VRE.EmpID = E.EmpID
        WHERE VRE.VR_RefID = @EntryID
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '');

    RETURN ISNULL(@ReturnStr, '');
END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VAccounts]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VAccounts];
END
GO

-- ------------------------------------------------------------------------------
-- 3. VIEWS DEFINITIONS (CREATE OR ALTER)
-- ------------------------------------------------------------------------------

CREATE   VIEW [dbo].[VAccounts]
AS
SELECT  dbo.Heads.Head AS MainHead
        ,dbo.Accounts.AccNo,dbo.Accounts.AccTitle,dbo.Accounts.Active,dbo.Accounts.Balance,dbo.Accounts.ComputerName,dbo.Accounts.CreateDT
        ,dbo.Accounts.openbal,dbo.Accounts.opendate,dbo.Accounts.Parent,dbo.Accounts.SubAccOf,dbo.Accounts.Type,dbo.Accounts.UserName
        ,dbo.Accounts.BudgetAccount,dbo.Accounts.BudgetAccount_Glance
        ,mainAccounts.AccTitle AS ParentAccTitle
        ,dbo.VAccountGroupsForExpenseReport.GroupName,dbo.VAccountGroupsForExpenseReport.EntryID AS GFER_EntryID
FROM dbo.Accounts 
INNER JOIN dbo.Heads ON dbo.Accounts.Type = dbo.Heads.code 
LEFT OUTER JOIN dbo.Accounts mainAccounts ON dbo.Accounts.SubAccOf = mainAccounts.AccNo
LEFT OUTER JOIN dbo.VAccountGroupsForExpenseReport ON dbo.Accounts.AccNo = dbo.VAccountGroupsForExpenseReport.AccNo;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VChqLedger]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VChqLedger];
END
GO



CREATE   VIEW [dbo].[VChqLedger]
AS

SELECT        dbo.Cheque.SNo, dbo.Cheque.CDate, dbo.Cheque.AccNo, dbo.Cheque.Description, dbo.Cheque.Amount, dbo.Cheque.BankID, 
dbo.Cheque.chequeno, dbo.Cheque.ChqBookNo, dbo.Cheque.chequeType, dbo.Cheque.chequeDate, 
dbo.Cheque.Posted, dbo.Cheque.Payment, dbo.Cheque.Bounced, dbo.Cheque.ChqIsDue, dbo.Cheque.Payee AS ChqPayee, dbo.Cheque.BankAccNo, dbo.Cheque.ClearanceDT, dbo.Cheque.ChqPrintingDone, dbo.BankList.Bank, 
dbo.BankList.Branch, dbo.BankAccounts.AccNo AS BankAccNo_old, Accounts_1.AccTitle AS BankAccountTitle, dbo.BankAccounts.Chqformat, Accounts_2.AccTitle, dbo.Makers.NICNo
,dbo.VoucherInfo.Handed_Over_To,dbo.VoucherInfo.VchrNo,Makers.CompanyName AS MakerPayeeName,
ISNULL(dbo.Cheque.Payee,Makers.CompanyName) AS Payee
,dbo.Vouchers.VDate,dbo.Vouchers.SNo AS Vouchers_SNo

FROM dbo.Cheque
LEFT JOIN dbo.ChqBooks ON dbo.Cheque.ChqBookNo = dbo.ChqBooks.ChqBookNo
LEFT JOIN dbo.Accounts AS Accounts_2 ON dbo.Cheque.AccNo=Accounts_2.AccNo
LEFT JOIN dbo.Makers ON dbo.Cheque.AccNo=dbo.Makers.AccNo
LEFT JOIN dbo.BankAccounts ON dbo.Cheque.BankAccNo = dbo.BankAccounts.AccNo
LEFT JOIN dbo.BankList ON dbo.BankList.BankID = dbo.BankAccounts.BankID
LEFT JOIN dbo.Accounts AS Accounts_1 ON dbo.BankAccounts.AccNo = Accounts_1.AccNo
LEFT JOIN dbo.Vouchers ON dbo.Cheque.BankAccNo=Vouchers.AccNo AND dbo.Cheque.SNo=dbo.Vouchers.CSNo
LEFT JOIN dbo.VoucherInfo ON Vouchers.VchrNO=VoucherInfo.VchrNo
                         

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VClosedLots_Simple_Main]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VClosedLots_Simple_Main];
END
GO

CREATE   VIEW [dbo].[VClosedLots_Simple_Main]
AS
SELECT        DISTINCT T1.ItemCode, T1.OrderNo, T1.LotNo
FROM            (SELECT        ItemCode, LotNo, OrderNo, Description, IssQty - RcvdQty AS Qty, ProcessID, VendID, ReWorkLot, DT, 0 AS LotType,0 AS VRD_EntryID
                          FROM            dbo.VVendIssdDetail_ForRunningLots
                          WHERE        (LotNo <> '0') AND (IssQty - RcvdQty > 0) AND (EntryID NOT IN
                                                        (SELECT        Issue_RefID
                                                          FROM            dbo.VendRcvdDetail))
                          UNION
                          SELECT        ItemCode, LotNo, OrderNo, Description, RcvdQty - IssQty - ISNULL(Wastage, 0) - ISNULL(ReWorkQty, 0) AS Qty, ProcessID, VendID, ReWorkLot, DT, 1 AS LotType,dbo.VVendRcvdDetail_Simple.EntryID AS VRD_EntryID
                          FROM            dbo.VVendRcvdDetail_Simple
                          WHERE        (LotNo <> '0') AND (RcvdQty - IssQty - ISNULL(Wastage, 0) - ISNULL(ReWorkQty, 0) > 0) AND (EntryID NOT IN
                                                       (SELECT        Rcvd_RefID
                                                         FROM            dbo.VendIssdDetail)) AND (ISNULL(Opening_RefID, 0) = 0)) AS T1 LEFT OUTER JOIN
                         dbo.VLotWithR2InDate ON T1.LotNo = dbo.VLotWithR2InDate.LotNo LEFT OUTER JOIN
                         dbo.VItemPolishingProcessSNo ON T1.ItemCode = dbo.VItemPolishingProcessSNo.ItemCode LEFT OUTER JOIN
                         dbo.ItemProcesses ON T1.ItemCode = dbo.ItemProcesses.ItemID AND dbo.ItemProcesses.ProcessID = T1.ProcessID
WHERE        (T1.LotNo IN
                             (SELECT        LotNo
                               FROM            dbo.Lots_Closed))
--ORDER BY T1.ProcessID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VCustomInvoiceDetail_From_PList_Total]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VCustomInvoiceDetail_From_PList_Total];
END
GO

CREATE   VIEW dbo.VCustomInvoiceDetail_From_PList_Total
AS
SELECT     CustomInvoice, OrderNo, CompItemID, SUM(Qty) AS TotalQty
FROM        dbo.VCustomInvoiceDetail_From_PList
GROUP BY CustomInvoice, OrderNo, CompItemID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VCustomPList_Itemwise]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VCustomPList_Itemwise];
END
GO

CREATE   VIEW [dbo].[VCustomPList_Itemwise]
AS
SELECT     TOP (100) PERCENT MIN(CartonFrom) AS FromCartonNo, MAX(CartonTo) AS ToCartonNo, CustomInvoice, OrderItemID, SUM(Qty) AS Expr1, COUNT(*) AS Tots
FROM        dbo.CustomPList
--WHERE     (CustomInvoice = '2322311')
GROUP BY CustomInvoice, OrderItemID
ORDER BY FromCartonNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VDispatchList]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VDispatchList];
END
GO


CREATE VIEW [dbo].[VDispatchList]
AS
SELECT dbo.DispatchList.DispatchListNo, dbo.DispatchList.DT, CASE WHEN NewFormat = 1 THEN T2.TotalCartons ELSE T1.TotalCartons END AS TotalCartons, dbo.DispatchList.EntryID, dbo.DispatchList.NewFormat, dbo.DispatchList.Finalyzed, dbo.DispatchList.FinalyzedDT, dbo.DispatchList.FinalyzedUserName, dbo.DispatchList.FinalyzedMachineName, dbo.DispatchList.CustCode, 
         T2.TotalQty
         ,dbo.DispatchList.UserName , dbo.DispatchList.MachineName , dbo.DispatchList.EntryDT , dbo.DispatchList.Country 
FROM  dbo.DispatchList LEFT OUTER JOIN
             (SELECT dbo.DispatchListDetails_Adv.RefID, MAX(dbo.DispatchListDetail_Cartons.CartonTo) AS TotalCartons
            FROM  dbo.DispatchListDetail_Cartons INNER JOIN
                     dbo.DispatchListDetails_Adv ON dbo.DispatchListDetails_Adv.EntryID = dbo.DispatchListDetail_Cartons.RefID
            GROUP BY dbo.DispatchListDetails_Adv.RefID) AS T1 ON dbo.DispatchList.EntryID = T1.RefID LEFT OUTER JOIN
             (SELECT DispatchListDetails_Adv_1.RefID, MAX(dbo.DispatchListDetail_Inners.CartonNo) AS TotalCartons, SUM(dbo.DispatchListDetail_Inners.Qty) AS TotalQty
            FROM  dbo.DispatchListDetail_Inners INNER JOIN
                     dbo.DispatchListDetails_Adv AS DispatchListDetails_Adv_1 ON DispatchListDetails_Adv_1.EntryID = dbo.DispatchListDetail_Inners.RefID
            GROUP BY DispatchListDetails_Adv_1.RefID) AS T2 ON dbo.DispatchList.EntryID = T2.RefID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VEmpOTHrs_New]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VEmpOTHrs_New];
END
GO

CREATE   VIEW [dbo].[VEmpOTHrs_New]
AS
SELECT     EmpID, DT, SUM(PayableHrs) AS ValidOTHrs
FROM         dbo.EmpTimes
WHERE     (OverTime = 1) AND  (DateName(WeekDay,DT)<>'Sunday')
AND DT NOT IN (select DT FROM Holidays)
GROUP BY EmpID, DT
UNION ALL
SELECT     EmpID, DT, SUM(PayableHrs) AS ValidOTHrs
FROM         dbo.EmpTimes
WHERE (DateName(WeekDay,DT)<>'Sunday')
AND DT IN (select DT FROM Holidays)
GROUP BY EmpID, DT

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VEmpProductivityReport_Production_New]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VEmpProductivityReport_Production_New];
END
GO

CREATE   VIEW [dbo].[VEmpProductivityReport_Production_New]
AS
SELECT dbo.Employees.empid, dbo.Employees.name, dbo.vendRcvdDetail.ItemCode, dbo.Items.ItemName, VEmpTimes2.DT, dbo.Processes.Description, dbo.VendReceived.ProcessID, ROUND(dbo.vendRcvdDetail.RcvdQty / TCount.EmpCount, 0) AS RcvdQty, 0 AS Type, CAST(CONVERT(VARCHAR(50), dbo.VendReceived.DT, 6) AS DATETIME) AS DTOnly, dbo.vendRcvdDetail.LotNo, 
         dbo.Departments.name AS Dname, TEmpTarget.Qty, dbo.VendReceived.OverTime, dbo.Employees.deptid, TEmpTargetOT.OTQty, TRegularGroupTargets.Qty AS GroupTargetQtyRegular, TOTGroupTargets.Qty AS GroupTargetQtyOverTime, TOTGroupTargets.GroupAvgOT, TRegularGroupTargets.GroupAvgRegular,
		 VEmpTimes2.FirstInTime,VEmpTimes2.SecondOutTime,ItemGroups.Description AS ItemGroupName
FROM VEmpTimes2 LEFT JOIN 
         dbo.VendReceived_Employees ON VEmpTimes2.empid=VendReceived_Employees.EmpID 
		 --dbo.VendReceived.EntryID = dbo.VendReceived_Employees.VR_RefID  LEFT JOIN--INNER JOIN
		 LEFT JOIN dbo.VendReceived ON dbo.VendReceived.EntryID = dbo.VendReceived_Employees.VR_RefID --AND VEmpTimes2.DT=VendReceived.DT  
		 LEFT JOIN --INNER JOIN

         dbo.Employees ON VEmpTimes2.EmpID = dbo.Employees.empid LEFT JOIN --INNER JOIN
         dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID LEFT JOIN --INNER JOIN
         dbo.Items ON dbo.vendRcvdDetail.ItemCode = dbo.Items.ItemID LEFT JOIN --INNER JOIN
         dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID LEFT JOIN --INNER JOIN
         dbo.Processes ON dbo.VendReceived.ProcessID = dbo.Processes.ProcessID LEFT JOIN --INNER JOIN
         dbo.Departments ON dbo.Employees.deptid = dbo.Departments.deptid LEFT OUTER JOIN
             (SELECT EmpID, ProcessID, GroupID, Qty
            FROM  dbo.EmpDailyTargets
            WHERE (OverTime = 0)) AS TEmpTarget 
			ON dbo.Processes.ProcessID = TEmpTarget.ProcessID AND dbo.Items.GroupID = TEmpTarget.GroupID AND dbo.Employees.empid = TEmpTarget.EmpID LEFT OUTER JOIN

             (SELECT EmpID, ProcessID, GroupID, Qty AS OTQty
            FROM  dbo.EmpDailyTargets AS EmpDailyTargets_1
            WHERE (OverTime = 1)) AS TEmpTargetOT ON dbo.Processes.ProcessID = TEmpTargetOT.ProcessID AND dbo.Items.GroupID = TEmpTargetOT.GroupID AND dbo.Employees.empid = TEmpTargetOT.EmpID LEFT OUTER JOIN
             (SELECT VR_RefID, COUNT(VR_RefID) AS EmpCount
            FROM  dbo.VendReceived_Employees AS VendReceived_Employees_1
            GROUP BY VR_RefID) AS TCount ON dbo.VendReceived.EntryID = TCount.VR_RefID LEFT OUTER JOIN
             (SELECT EmpID, SUM(Qty) AS Qty, AVG(Qty) AS GroupAvgRegular
            FROM  dbo.EmpDailyTargets AS EmpDailyTargets_Group
            WHERE (OverTime = 0)
            GROUP BY EmpID) AS TRegularGroupTargets ON dbo.Employees.empid = TRegularGroupTargets.EmpID LEFT OUTER JOIN
             (SELECT EmpID, SUM(Qty) AS Qty, AVG(Qty) AS GroupAvgOT
            FROM  dbo.EmpDailyTargets AS EmpDailyTargets_GroupOT
            WHERE (OverTime = 1)
            GROUP BY EmpID) AS TOTGroupTargets ON dbo.Employees.empid = TOTGroupTargets.EmpID
--WHERE VEmpTimes2.DT=VendReceived.DT

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VEmpTimes2]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VEmpTimes2];
END
GO


CREATE VIEW [dbo].[VEmpTimes2]
AS
SELECT     TOP 100 PERCENT dbo.Employees.empid, dbo.Employees.deptid, dbo.Employees.name, dbo.Employees.fname, dbo.Employees.Rel, dbo.Employees.Active, 
                      dbo.VEmpDailyEntries.DT, EmpTimes_1.InTime AS FirstInTime, EmpTimes_1.OutTime AS FirstOutTime, EmpTimes_2.InTime AS SecondInTime, 
                      EmpTimes_2.OutTime AS SecondOutTime, EmpTimes_3.InTime AS OTInTime, EmpTimes_3.OutTime AS OTOutTime, dbo.VEmpDailyEntries.FirstEntryID, 
                      ISNULL(EmpTimes_1.Hrs, 0) + ISNULL(EmpTimes_2.Hrs, 0) AS Hrs, ISNULL(EmpTimes_1.PayableHrs, 0) + ISNULL(EmpTimes_2.PayableHrs, 0) AS PayableHrs, 
                      ISNULL(EmpTimes_3.Hrs, 0) AS OTHrs, dbo.VEmpDailyEntries.TotalEntries, dbo.VEmpDailyEntries.SecondEntryID, dbo.VEmpDailyEntries.OverTimeEntryID, 
                      dbo.VEmpSettings.LunchInTime, dbo.VEmpSettings.LunchOutTime,AttendanceSheet.LateHours
FROM         dbo.VEmpDailyEntries INNER JOIN
                      dbo.Employees ON dbo.VEmpDailyEntries.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                      dbo.VEmpSettings ON dbo.Employees.empid = dbo.VEmpSettings.empid LEFT OUTER JOIN
                      dbo.VEmpPayableHrs EmpTimes_1 ON dbo.VEmpDailyEntries.FirstEntryID = EmpTimes_1.EntryID LEFT OUTER JOIN
                      dbo.VEmpPayableHrs EmpTimes_2 ON dbo.VEmpDailyEntries.SecondEntryID = EmpTimes_2.EntryID LEFT OUTER JOIN
                      dbo.VEmpPayableHrsOverTime EmpTimes_3 ON dbo.VEmpDailyEntries.OverTimeEntryID = EmpTimes_3.EntryID
					  INNER JOIN AttendanceSheet ON EmpTimes_1.EmpID=AttendanceSheet.EmpID AND EmpTimes_1.DT=AttendanceSheet.DT
ORDER BY dbo.Employees.empid


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VFCustomerCatalog_Ex]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VFCustomerCatalog_Ex];
END
GO






CREATE   VIEW [dbo].[VFCustomerCatalog_Ex]
AS
SELECT     dbo.FCustomerCatalog.EntryID, dbo.FCustomerCatalog.CustCode, dbo.FCustomerCatalog.Country, dbo.FCustomerCatalog.ItemID, dbo.FCustomerCatalog.Unit, 
                      dbo.FCustomerCatalog.CompItemID, dbo.FCustomerCatalog.Description, dbo.FCustomerCatalog.FOB, dbo.FCustomerCatalog.CnFAir, dbo.FCustomerCatalog.CnFSea, 
                      dbo.FCustomerCatalog.CIFAir, dbo.FCustomerCatalog.CIFSea, dbo.FCustomerCatalog.PackingMode, dbo.VItems.ItemName AS CompDesc, dbo.VItems.CatID, 
                      dbo.ItemCatagories.Description AS CatDesc, dbo.VItems.FOB AS CustomFOB, dbo.VItems.CIFSea AS CustomCIFSea, dbo.VItems.CIFAir AS CustomCIFAir, 
                      dbo.VItems.CnFSea AS CustomCnFSea, dbo.VItems.CnFAir AS CustomCnFAir, dbo.FCustomerCatalog.ExWorks, dbo.FCustomerCatalog.CnIAir, 
                      dbo.FCustomerCatalog.CnISea, dbo.FCustomerCatalog.ExWorks AS CustomExWorks, dbo.FCustomerCatalog.CnIAir AS CustomCnIAir, 
                      dbo.FCustomerCatalog.CnISea AS CustomCnISea, dbo.FCustomerCatalog.FOBTop, dbo.FCustomerCatalog.ExWorksTop, dbo.FCustomerCatalog.CnFAirTop, 
                      dbo.FCustomerCatalog.CnFSeaTop, dbo.FCustomerCatalog.CIFAirTop, dbo.FCustomerCatalog.CIFSeaTop, dbo.FCustomerCatalog.CnIAirTop, 
                      dbo.FCustomerCatalog.CnISeaTop, dbo.FCustomerCatalog.ILO, dbo.FCustomerCatalog.OCR, dbo.FCustomerCatalog.OCR2, dbo.FCustomerCatalog.FinQuality, 
                      dbo.FCustomerCatalog.StampInstructions, dbo.FCustomerCatalog.SpecialInstructions, dbo.VItems.InActive
					  ,dbo.VItems.ItemSize,dbo.VItems.SizeUnit,dbo.VItems.TipSize
					  ,TRL.Running_Lots_No,TRL.Running_Lots_Qty,TForging.QtyInStock AS Forging_Stock
					  ,dbo.VItems.ItemGroup,dbo.VItems.GroupID
                      ,dbo.FCustomerCatalog.BarcodeNo,dbo.FCustomerCatalog.Temper_Rate,dbo.FCustomerCatalog.PackingInstructions,dbo.FCustomerCatalog.ItemColor
                      ,dbo.FCustomerCatalog.BarcodeFile,dbo.FCustomerCatalog.First_Inspection_Rate,dbo.FCustomerCatalog.MDMA,dbo.FCustomerCatalog.SFDA_Listing_No
                      ,dbo.FCustomerCatalog.MD_Group
FROM         dbo.FCustomerCatalog INNER JOIN
                      dbo.VItems ON dbo.FCustomerCatalog.CompItemID = dbo.VItems.ItemID INNER JOIN
                      dbo.ItemCatagories ON dbo.VItems.CatID = dbo.ItemCatagories.CatID
					  LEFT JOIN (SELECT ItemCode,COUNT(LotNo) AS Running_Lots_No,SUM(Qty) AS Running_Lots_Qty FROM VRunningLots_Simple GROUP BY ItemCode) TRL ON dbo.VItems.ItemID=TRL.ItemCode
					  LEFT JOIN (SELECT ItemID,SUM(VRM.QtyInStock) AS QtyInStock FROM VRM
									INNER JOIN ItemsRMComp ON VRM.RMID=ItemsRMComp.RMID
									INNER JOIN RMGroupIDsForForging ON VRM.GroupID=RMGroupIDsForForging.Group_ID GROUP BY ItemID) TForging ON dbo.VItems.ItemID=TForging.ItemID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VFCustomers]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VFCustomers];
END
GO


CREATE VIEW [dbo].[VFCustomers]
AS
SELECT        dbo.ForeignCustomers.CustCode, dbo.ForeignCustomers.Country, dbo.ForeignCustomers.Name, dbo.ForeignCustomers.Address, dbo.ForeignCustomers.FakeAddress, dbo.ForeignCustomers.Phone1, 
                         dbo.ForeignCustomers.Phone2, dbo.ForeignCustomers.Phone3, dbo.ForeignCustomers.Email1, dbo.ForeignCustomers.Email2, dbo.ForeignCustomers.Fax1, dbo.ForeignCustomers.Fax2, dbo.ForeignCustomers.Fax3, 
                         dbo.ForeignCustomers.URL, dbo.ForeignCustomers.Curr, dbo.ForeignCustomers.Cont1name, dbo.ForeignCustomers.cont1phone, dbo.ForeignCustomers.cont1email, dbo.ForeignCustomers.cont1Mobile, 
                         dbo.ForeignCustomers.Cont1Skype, dbo.ForeignCustomers.cont2name, dbo.ForeignCustomers.cont2phone, dbo.ForeignCustomers.cont2email, dbo.ForeignCustomers.cont2Mobile, dbo.ForeignCustomers.Cont2Skype, 
                         dbo.ForeignCustomers.AgentName, dbo.ForeignCustomers.AgentAddress, dbo.ForeignCustomers.AgentPhone1, dbo.ForeignCustomers.AgentPhone2, dbo.ForeignCustomers.AgentFax1, dbo.ForeignCustomers.AgentFax2, 
                         dbo.ForeignCustomers.AgentEmail, dbo.ForeignCustomers.AgentWeb, dbo.ForeignCustomers.ACName, dbo.ForeignCustomers.ACPhone, dbo.ForeignCustomers.ACMobile, dbo.ForeignCustomers.AgentURL, 
                         dbo.ForeignCustomers.ACEmail, dbo.ForeignCustomers.DTFormat, dbo.ForeignCustomers.Active, dbo.ForeignCustomers.AccNo, dbo.ForeignCustomers.NearestAirport, dbo.ForeignCustomers.NearestRailwaystation, 
                         dbo.ForeignCustomers.FedexNo, dbo.ForeignCustomers.DHLNo, dbo.ForeignCustomers.TradeTerms, dbo.ForeignCustomers.PaymentTerms, dbo.ForeignCustomers.AcceptsExtraQty, dbo.ForeignCustomers.SpecialInstructions, 
                         dbo.ForeignCustomers.City, dbo.ForeignCustomers.MaxDiscount, dbo.ForeignCustomers.PaymentDays, dbo.ForeignCustomers.LateOrderAlerts, dbo.ForeignCustomers.Customer_Type, 
                         dbo.ForeignCustomers.OuterPackingLabel, dbo.ForeignCustomers.InnerPackingLabel, dbo.ForeignCustomers.ShowCustomerRef, dbo.ForeignCustomers.FinishingQuality, dbo.ForeignCustomers.Stamps, 
                         dbo.ForeignCustomers.RcvdVia, dbo.ForeignCustomers.DefaultPort, dbo.ForeignCustomers.DefaultShipMethod, dbo.ForeignCustomers.Cont1Designation, dbo.ForeignCustomers.Cont2Designation, 
                         dbo.ForeignCustomers.CustomerSource, dbo.ForeignCustomers.Name AS AccTitle, dbo.VPorts.CityID AS DisCityID, dbo.VPorts.City AS DisCity, dbo.VPorts.PortID AS DisPortID, dbo.VPorts.Port AS DisPort, 
                         dbo.ForeignCustomers.FC_Note_I, dbo.ForeignCustomers.FC_Note_II, dbo.ForeignCustomers.Inner_Label_Manual_I, dbo.ForeignCustomers.Inner_Label_Manual_II
FROM            dbo.ForeignCustomers LEFT OUTER JOIN
                         dbo.VPorts ON dbo.ForeignCustomers.DefaultPort = dbo.VPorts.PortID


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VFOrderItems_Lots]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VFOrderItems_Lots];
END
GO

CREATE   VIEW [dbo].[VFOrderItems_Lots]
AS
SELECT        CustCode, Country, OrderNo, InternalRefNo, CompItemCode, DeliveryDT, Qty, ShippedQty, DeliveryStatus, ItemName, ID, DT, ItemCode, Description, Packaging, OrderDeliveryDT, Quality, ItemQuality, GroupID, Remarks, SortNo, 
                         Price, Item_Finishing_Type, dbo.RunningLots(OrderNo, CompItemCode) AS Lot_Processes
FROM            dbo.VFOrderItems

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VFOrderItems_PTC]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VFOrderItems_PTC];
END
GO

-- ====================================================================================================
-- PART 5: VIEWS
-- ====================================================================================================
CREATE   VIEW dbo.VFOrderItems_PTC
AS
SELECT DISTINCT dbo.FOrderItems.OrderNo, dbo.FOrderItems.ItemCode, T2.PTC
FROM        dbo.FOrderItems LEFT OUTER JOIN
                      (SELECT     dbo.vendRcvdDetail.OrderNo, dbo.vendRcvdDetail.ItemCode, dbo.vendRcvdDetail.LotNo + ' (' + CAST(dbo.vendRcvdDetail.RcvdQty AS VARCHAR(5)) + ')' AS PTC
                       FROM        dbo.vendRcvdDetail WITH (NOLOCK) INNER JOIN
                                             (SELECT     MAX(EntryID) AS MAXEntryID
                                              FROM        dbo.vendRcvdDetail WITH (NOLOCK)
                                              GROUP BY LotNo) AS T1 ON dbo.vendRcvdDetail.EntryID = T1.MAXEntryID) AS T2 ON dbo.FOrderItems.OrderNo = T2.OrderNo AND dbo.FOrderItems.ItemCode = T2.ItemCode

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VFOrderList]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VFOrderList];
END
GO

-- -------------------------------------------------------------
-- Update VFOrderList to include TotalAuthorizedArticles & TotalPendingArticles
-- -------------------------------------------------------------
CREATE   VIEW [dbo].[VFOrderList]
AS
SELECT     dbo.FCustomerOrders.OrderNo, dbo.FCustomerOrders.DT, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.F_OrderAmt(dbo.FCustomerOrders.OrderNo, 0) AS OrderAmt, 
                      dbo.ForeignCustomers.Curr, dbo.VFOrderTotalQty.TotalInvQty, dbo.VFOrderTotalQty.TotalOrderQty, dbo.FCustomerOrders.DeliveryDT, dbo.FCustomerOrders.CompanyRefID, 
                      dbo.Companies.CompanyName, dbo.VUnshippedOrderList.OrderNo AS UnshippedOrderNo, TotalShipped.TotalShippedQty, dbo.VCurrencyExchangeRates.ExchRate, 
                      dbo.FCustomerOrders.InternalRefNo, ROUND(dbo.VFOrderTotalQty.TotalWeight, 2) AS TotalWeight, TotalShipped.TotalShippedAmt, ISNULL(dbo.FCustomerFinalOrders.Cancelled, 0) AS Cancelled, 
                      TotalQuantities.TotalArticles, dbo.FCustomerFinalOrders.Remarks, TotalShipped.ShippedItemCount, TotalShipped.ShippedItemQty, TotalShipped.PartiallyShippedItemCount, 
                      TotalShipped.PartiallyShippedItemQty, TotalShipped.TotalBalanceQty, dbo.FCustomerOrders.OrderType, T1.TotalPlannedQty, dbo.FCustomerOrders.OrderPlanApproved
                      ,dbo.ForeignCustomers.LateOrderAlerts
                      ,ISNULL(dbo.FCustomerOrders.Authorized, 0) AS Authorized
                      ,dbo.FCustomerOrders.AuthorizedBy
                      ,dbo.FCustomerOrders.AuthorizedDT
                      ,ISNULL(TotalQuantities.TotalAuthorizedArticles, 0) AS TotalAuthorizedArticles
                      ,ISNULL(TotalQuantities.TotalPendingArticles, 0) AS TotalPendingArticles
FROM         dbo.FCustomerOrders INNER JOIN
                      dbo.ForeignCustomers ON dbo.FCustomerOrders.CustCode = dbo.ForeignCustomers.CustCode AND dbo.FCustomerOrders.Country = dbo.ForeignCustomers.Country INNER JOIN
                      dbo.VFOrderTotalQty ON dbo.FCustomerOrders.OrderNo = dbo.VFOrderTotalQty.OrderNo INNER JOIN
                      dbo.Companies ON dbo.FCustomerOrders.CompanyRefID = dbo.Companies.EntryID LEFT OUTER JOIN
                      dbo.VUnshippedOrderList ON dbo.FCustomerOrders.OrderNo = dbo.VUnshippedOrderList.OrderNo INNER JOIN
                          (SELECT     OrderNo, 
                                      SUM(Qty) AS TotalOrderQty, 
                                      COUNT(*) AS TotalArticles,
                                      SUM(CASE WHEN ISNULL(Authorized, 0) = 1 AND ISNULL(AuthorizedQty, 0) >= ISNULL(Qty, 0) THEN 1 ELSE 0 END) AS TotalAuthorizedArticles,
                                      SUM(CASE WHEN ISNULL(Authorized, 0) = 0 OR ISNULL(AuthorizedQty, 0) < ISNULL(Qty, 0) THEN 1 ELSE 0 END) AS TotalPendingArticles
                            FROM          dbo.FOrderItems
                            GROUP BY OrderNo) AS TotalQuantities ON dbo.FCustomerOrders.OrderNo = TotalQuantities.OrderNo LEFT OUTER JOIN
                      dbo.FCustomerFinalOrders ON dbo.FCustomerOrders.OrderNo = dbo.FCustomerFinalOrders.OrderNo LEFT OUTER JOIN
                          (SELECT     OrderNo, SUM(ShippedQty) AS TotalShippedQty, SUM(ShippedQty * Price) AS TotalShippedAmt, SUM(CASE WHEN ShippedQty >= Qty THEN 1 ELSE 0 END) AS ShippedItemCount, 
                                                   SUM(CASE WHEN ShippedQty >= Qty THEN ShippedQty ELSE 0 END) AS ShippedItemQty, SUM(CASE WHEN ShippedQty > 0 AND ShippedQty < Qty THEN 1 ELSE 0 END) 
                                                   AS PartiallyShippedItemCount, SUM(CASE WHEN ShippedQty > 0 AND ShippedQty < Qty THEN ShippedQty ELSE 0 END) AS PartiallyShippedItemQty, 
                                                   SUM(CASE WHEN ShippedQty >= Qty THEN 0 ELSE Qty - ShippedQty END) AS TotalBalanceQty
                            FROM          dbo.VFOrderItemswithShippedQty
                            WHERE      (CompItemCode IN
                                                       (SELECT     ItemID
                                                         FROM          dbo.Items))
                            GROUP BY OrderNo) AS TotalShipped ON dbo.FCustomerOrders.OrderNo = TotalShipped.OrderNo LEFT OUTER JOIN
                      dbo.VCurrencyExchangeRates ON dbo.ForeignCustomers.Curr = dbo.VCurrencyExchangeRates.Currency LEFT OUTER JOIN
                          (SELECT     OrderNo, SUM(Qty) AS TotalPlannedQty
                            FROM          dbo.OrderPlanningDetails
                            GROUP BY OrderNo) AS T1 ON dbo.FCustomerOrders.OrderNo = T1.OrderNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER OFF;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VGatePass_Return_Details]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VGatePass_Return_Details];
END
GO


CREATE     VIEW [dbo].[VGatePass_Return_Details]
AS
SELECT     dbo.GatePass_Return_Details.EntryID, dbo.GatePass_Return_Details.MEGO_RefID, dbo.GatePass_Return_Details.Detail, 
                      dbo.GatePass_Return_Details.Unit, dbo.GatePass_Return_Details.Packing, dbo.GatePass_Return_Details.Qty, ISNULL(T1.QtyRcvd, 0) 
                      AS TotalQtyRcvd
FROM         dbo.GatePass_Return_Details LEFT OUTER JOIN
                          (SELECT     ME_GO_D_RefID, SUM(QtyRcvd) AS QtyRcvd
                             FROM         dbo.GatePass_Return_Receiving_Detail
                             GROUP BY ME_GO_D_RefID) AS T1 ON dbo.GatePass_Return_Details.EntryID = T1.ME_GO_D_RefID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VGateRcvdByPO]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VGateRcvdByPO];
END
GO

CREATE   VIEW dbo.VGateRcvdByPO
AS
SELECT PORefNo, WORef, MaterialID, SUM(ISNULL(QtyRcvd, 0) - ISNULL(QtyReturned, 0)) AS TotalGateRcvd
FROM dbo.VendRcvdDetailPO
WHERE (ISNULL(Checked, 0) = 0)
GROUP BY PORefNo, WORef, MaterialID;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VGateRcvdByWO]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VGateRcvdByWO];
END
GO

CREATE   VIEW dbo.VGateRcvdByWO
AS
SELECT WORef, MaterialID, SUM(ISNULL(QtyRcvd, 0) - ISNULL(QtyReturned, 0)) AS TotalGateRcvd
FROM dbo.VendRcvdDetailPO
WHERE (ISNULL(Checked, 0) = 0)
GROUP BY WORef, MaterialID;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VHubSupervisors]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VHubSupervisors];
END
GO

-- ====================================================================================================
-- [45/49] SCRIPT: 26-09-26-15-VHubSupervisors.sql
-- ====================================================================================================
CREATE   VIEW [dbo].[VHubSupervisors]
AS
SELECT 
    s1.GroupID,
    s1.Hub_Name,
    STUFF((
        SELECT ', ' + s2.UserName
        FROM dbo.ProcessGroup_Hub_Supervisors s2
        WHERE s2.GroupID = s1.GroupID AND s2.Hub_Name = s1.Hub_Name
        ORDER BY s2.UserName
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS Supervisors
FROM dbo.ProcessGroup_Hub_Supervisors s1
GROUP BY s1.GroupID, s1.Hub_Name

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItem_RunningLots]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItem_RunningLots];
END
GO

CREATE   VIEW dbo.VItem_RunningLots
AS
SELECT     ItemCode, COUNT(*) AS RunningLots, SUM(Qty) AS Qty
FROM        dbo.VRunningLots
GROUP BY ItemCode

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItemOpenPOs_itemwise]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItemOpenPOs_itemwise];
END
GO

CREATE   VIEW [dbo].[VItemOpenPOs_itemwise]
AS
SELECT ItemCode,SUM(Qty) AS OpenPOQty FROM VItemOpenPOs GROUP BY ItemCode

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItems]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItems];
END
GO



CREATE   VIEW [dbo].[VItems]
AS
SELECT dbo.Items.CreateDT, dbo.Items.CatID, dbo.Items.ItemID, dbo.Items.ItemName, dbo.Items.Unit, dbo.Items.Type, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.ItemUsage, dbo.Items.Attributes, dbo.Items.FinQuality, dbo.Items.SteelUsed, dbo.Items.Gage, dbo.Items.PacknLabel, dbo.Items.AssetAccNo, dbo.Items.SaleAccNo, dbo.Items.SRTAccNo, dbo.Items.FOB, dbo.Items.CIFSea, 
         dbo.Items.CIFAir, dbo.Items.CnFSea, dbo.Items.CnFAir, dbo.Items.ExWorks, dbo.Items.CnISea, dbo.Items.CnIAir, dbo.Items.ItemCurr, dbo.Items.UnitWeight, dbo.Items.ForgingWeight, dbo.Items.WasteVisible, dbo.Items.FinishedWeight, dbo.Items.InActive, dbo.Items.ItemPic, dbo.Items.RevID, dbo.Items.GroupID, dbo.Items.PlantRate, dbo.Items.SnaffRate, dbo.Items.StampRate, 
         dbo.Items.ReorderPoint, dbo.Items.FinishDescription1, dbo.Items.FinishDescription2, dbo.Items.MakerDescription1, dbo.Items.MakerDescription2, dbo.Items.Tagging, dbo.Items.EAN128, dbo.Items.CustomDescription, dbo.Items.ItemColor, dbo.Items.PackingInstructions, dbo.Items.MasterCartonL, dbo.Items.MasterCartonW, dbo.Items.MasterCartonH, dbo.Items.SmallCartonL, 
         dbo.Items.SmallCartonW, dbo.Items.SmallCartonH, dbo.Items.PolyBag, dbo.Items.FixedPackingUnit, dbo.VItemWithStocks.TotalFinishedStock AS InHand, dbo.Items.OpenBal, dbo.Items.ItemNameUrdu, dbo.Items.PriceForCost, dbo.Items.TechnicalDrawing, dbo.Items.TipSize, dbo.Items.MinLevel, dbo.Items.MaxLevel, dbo.Items.ReOrderLevel, dbo.Items.ItemType, dbo.Items.FOBTop, 
         dbo.Items.CIFSeaTop, dbo.Items.CIFAirTop, dbo.Items.CnFSeaTop, dbo.Items.CnFAirTop, dbo.Items.ExWorksTop, dbo.Items.CnISeaTop, dbo.Items.CnIAirTop, dbo.Items.GTINBarcodeNo, dbo.Items.Description, dbo.Items.UMDNSCode, dbo.Items.FDAListingNo, dbo.Items.FDAProductCode, dbo.Items.EuropeanRegNo, dbo.Items.FDA510K, dbo.Items.ItemMaxLotSize, 
         dbo.Items.ItemLotSizeBuffer, dbo.Items.ReadyFinishPrice, dbo.Items.FillingPrice, dbo.Items.MainGroupID, dbo.Items.Sample, dbo.Items.Type AS Typ1, Accounts_3.AccTitle AS AssetAccNoText, Accounts_1.AccTitle AS SaleAccNoText, Accounts_2.AccTitle AS SRTAccNoText, dbo.SteelGages.Gage AS SteelGage, dbo.SteelTypes.SteelType, dbo.ItemGroups.Description AS ItemGroup, 
         dbo.ItemGroups.GrpColor, dbo.ItemCatagories.Description AS Category, dbo.ItemGroupsMain.MainGroupName, dbo.Items.SmallBoxPcs, dbo.Items.MasterCartonSmallBoxes, dbo.Items.AvailableForECommerce, dbo.Items.PicForECommerce, dbo.Items.DescriptionForECommerce, dbo.Items.PriceForECommerce, dbo.Items.GroupIDForECommerce, 
         dbo.ItemGroups_ECommerce.ECommerceGroupName, dbo.Items.POInustructions, dbo.Items.SFDA_Name, dbo.Items.SFDA_No, dbo.Items.SFDA_Listing, dbo.Items.HRC_From, dbo.Items.HRC_To,
         dbo.Items.Fig_No, dbo.Items.Scale, dbo.Items.Additional_Detail, dbo.SteelGages.GageUnit, dbo.Items.MakerDescription
         ,0 AS Parts
         ,dbo.Items.MasterCartonTotalPcs,dbo.Items.TipCap
FROM  dbo.SteelTypes RIGHT OUTER JOIN
         dbo.Items INNER JOIN
         dbo.ItemCatagories ON dbo.Items.CatID = dbo.ItemCatagories.CatID LEFT OUTER JOIN
         dbo.ItemGroups_ECommerce ON dbo.Items.GroupIDForECommerce = dbo.ItemGroups_ECommerce.ID LEFT OUTER JOIN
         dbo.ItemGroupsMain ON dbo.Items.MainGroupID = dbo.ItemGroupsMain.MainGroupID LEFT OUTER JOIN
         dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID ON dbo.SteelTypes.SteelID = dbo.Items.SteelUsed LEFT OUTER JOIN
         dbo.SteelGages ON dbo.Items.Gage = dbo.SteelGages.GageID LEFT OUTER JOIN
         dbo.Accounts AS Accounts_2 ON dbo.Items.SRTAccNo = Accounts_2.AccNo LEFT OUTER JOIN
         dbo.Accounts AS Accounts_1 ON dbo.Items.SaleAccNo = Accounts_1.AccNo LEFT OUTER JOIN
         dbo.Accounts AS Accounts_3 ON dbo.Items.AssetAccNo = Accounts_3.AccNo LEFT OUTER JOIN
         dbo.VItemWithStocks ON dbo.Items.ItemID = dbo.VItemWithStocks.ItemID


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItems_Parts]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItems_Parts];
END
GO
CREATE VIEW [dbo].[VItems_Parts]
AS
SELECT dbo.Items_Parts.EntryID, dbo.Items_Parts.ItemID, dbo.Items_Parts.ItemID_Part, dbo.Items_Parts.Qty, dbo.Items_Parts.ProcessID, dbo.Items.ItemName, T1.ItemName AS ItemName_Part, dbo.Processes.Description
FROM  dbo.Items_Parts INNER JOIN
         dbo.Items ON dbo.Items.ItemID = dbo.Items_Parts.ItemID INNER JOIN
         dbo.Items AS T1 ON dbo.Items_Parts.ItemID_Part = T1.ItemID INNER JOIN
         dbo.Processes ON dbo.Items_Parts.ProcessID = dbo.Processes.ProcessID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItems_StockReport]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItems_StockReport];
END
GO




CREATE VIEW [dbo].[VItems_StockReport]
AS
SELECT        dbo.VItems.CatID, dbo.VItems.ItemID, dbo.VItems.ItemName, dbo.VItems.Unit, dbo.VItems.Type, dbo.VItems.ItemSize, dbo.VItems.SizeUnit, dbo.VItems.ItemUsage, dbo.VItems.FinQuality, dbo.VItems.ItemGroup, 
                         dbo.VItems.InActive, TSF.SFStock, TForging.ForgingStock, dbo.VItems.MinLevel, dbo.VItems.MaxLevel, dbo.VItems.ReOrderLevel, TFinished.InHand, TRunning.InProductionQty, dbo.VStoreShelfs.ShelfNo, TFinished.Shelf_RefID, 
dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.StoreName, dbo.VItems.PriceForCost, dbo.VItems.FillingPrice,
ISNULL(TMakerPO.MaxRate,0) AS ItemMaxRate
,VItems.Parts
FROM            dbo.VStoreShelfs RIGHT OUTER JOIN
                             (SELECT        ItemID, SUM(NetQty) AS InHand, MAX(Shelf_RefID) AS Shelf_RefID
                                FROM            dbo.VItemsWithShelfWiseStock_Simple
                                GROUP BY ItemID) AS TFinished ON dbo.VStoreShelfs.EntryID = TFinished.Shelf_RefID RIGHT OUTER JOIN
                         dbo.VItems LEFT OUTER JOIN
                             (SELECT        ItemID, SUM(Qty) AS SFStock
                                FROM            dbo.VItemSFStock
                                GROUP BY ItemID) AS TSF ON dbo.VItems.ItemID = TSF.ItemID LEFT OUTER JOIN
                             (SELECT        dbo.ItemsRMComp.ItemID, SUM(dbo.VRM.QtyInStock) AS ForgingStock
                                FROM            dbo.ItemsRMComp INNER JOIN
                                                         dbo.VRM ON dbo.VRM.RMID = dbo.ItemsRMComp.RMID
                                GROUP BY dbo.ItemsRMComp.ItemID) AS TForging ON dbo.VItems.ItemID = TForging.ItemID ON TFinished.ItemID = dbo.VItems.ItemID LEFT OUTER JOIN
                             (SELECT        dbo.VRunningLots_Simple_PHF.ItemCode, SUM(dbo.VRunningLots_Simple_PHF.Qty) AS InProductionQty
                                FROM            dbo.VRunningLots_Simple_PHF INNER JOIN
                                                         dbo.FCustomerOrders ON dbo.VRunningLots_Simple_PHF.OrderNo = dbo.FCustomerOrders.OrderNo
                                WHERE        (dbo.FCustomerOrders.CustCode IN ('Stock'))
                                GROUP BY dbo.VRunningLots_Simple_PHF.ItemCode) AS TRunning ON dbo.VItems.ItemID = TRunning.ItemCode
LEFT JOIN (SELECT ItemCode,MAX(IssRate) As MaxRate FROM VVendRcvdDetail GROUP BY ItemCode) TMakerPO ON 
VItems.ItemID=TMakerPO.ItemCode

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItems_StockReport_New]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItems_StockReport_New];
END
GO

CREATE   VIEW [dbo].[VItems_StockReport_New]
AS
SELECT     dbo.VItems.CatID, dbo.VItems.ItemID, dbo.VItems.ItemName, dbo.VItems.Unit, dbo.VItems.Type, dbo.VItems.ItemSize, dbo.VItems.SizeUnit, dbo.VItems.ItemUsage, dbo.VItems.FinQuality, dbo.VItems.ItemGroup, dbo.VItems.InActive, TSF.SFStock, 
                  TForging.ForgingStock, dbo.VItems.MinLevel, dbo.VItems.MaxLevel, dbo.VItems.ReOrderLevel, TFinished.InHand, TRunning.InProductionQty, dbo.VStoreShelfs.ShelfNo, TFinished.Shelf_RefID, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.StoreName, 
                  dbo.VItems.PriceForCost, dbo.VItems.FillingPrice, dbo.VItems.GroupID
FROM        dbo.VStoreShelfs RIGHT OUTER JOIN
                      (SELECT     ItemID, SUM(NetQty) AS InHand, MAX(Shelf_RefID) AS Shelf_RefID
                       FROM        dbo.VItemsWithShelfWiseStock_Simple
                       GROUP BY ItemID) AS TFinished ON dbo.VStoreShelfs.EntryID = TFinished.Shelf_RefID RIGHT OUTER JOIN
                  dbo.VItems LEFT OUTER JOIN
                      (SELECT     ItemID, SUM(Qty) AS SFStock
                       FROM        dbo.VItemSFStock
                       GROUP BY ItemID) AS TSF ON dbo.VItems.ItemID = TSF.ItemID LEFT OUTER JOIN
                      (SELECT     dbo.ItemsRMComp.ItemID, SUM(dbo.VRM.QtyInStock) AS ForgingStock
                       FROM        dbo.ItemsRMComp INNER JOIN
                                         dbo.VRM ON dbo.VRM.RMID = dbo.ItemsRMComp.RMID
                       GROUP BY dbo.ItemsRMComp.ItemID) AS TForging ON dbo.VItems.ItemID = TForging.ItemID ON TFinished.ItemID = dbo.VItems.ItemID LEFT OUTER JOIN
                      (SELECT     dbo.VRunningLots_Simple_PHF.ItemCode, SUM(dbo.VRunningLots_Simple_PHF.Qty) AS InProductionQty
                       FROM        dbo.VRunningLots_Simple_PHF INNER JOIN
                                         dbo.FCustomerOrders ON dbo.VRunningLots_Simple_PHF.OrderNo = dbo.FCustomerOrders.OrderNo
                       WHERE     (dbo.FCustomerOrders.CustCode IN ('Stock'))
                       GROUP BY dbo.VRunningLots_Simple_PHF.ItemCode) AS TRunning ON dbo.VItems.ItemID = TRunning.ItemCode

WHERE TFinished.InHand>0

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItemsRMComp]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItemsRMComp];
END
GO

CREATE VIEW [dbo].[VItemsRMComp]
AS
SELECT        dbo.ItemsRMComp.EntryID, dbo.ItemsRMComp.ItemID, dbo.ItemsRMComp.RMID, dbo.ItemsRMComp.Qty, dbo.ItemsRMComp.ProcessID, dbo.VRM.RMName AS Description, dbo.Processes.Description AS ProcDesc, 
                         dbo.VRM.Description AS GroupDescription, dbo.Items.ItemName, dbo.VRM.RMID1, dbo.VRM.GroupID, dbo.VRM.QtyInStock,dbo.ItemsRMComp.Functional_Status
                         ,VRM.GroupName,VRM.Unit,VRM.InActive
FROM            dbo.ItemsRMComp INNER JOIN
                         dbo.VRM ON dbo.ItemsRMComp.RMID = dbo.VRM.RMID INNER JOIN
                         dbo.Items ON dbo.ItemsRMComp.ItemID = dbo.Items.ItemID LEFT OUTER JOIN
                         dbo.Processes ON dbo.ItemsRMComp.ProcessID = dbo.Processes.ProcessID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItemsRMCompWithLocationWiseStatus]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItemsRMCompWithLocationWiseStatus];
END
GO

CREATE   VIEW [dbo].[VItemsRMCompWithLocationWiseStatus]
AS
SELECT        dbo.ItemsRMComp.EntryID, dbo.ItemsRMComp.ItemID, dbo.ItemsRMComp.RMID, dbo.ItemsRMComp.Qty, dbo.ItemsRMComp.ProcessID, dbo.VRM.RMName AS Description, dbo.Processes.Description AS ProcDesc, 
                         dbo.VRM.Description AS GroupDescription, dbo.Items.ItemName, dbo.VRM.RMID1, dbo.VRM.GroupID, 
						 dbo.VRM.QtyInStock,dbo.ItemsRMComp.Functional_Status,
						 VMaterialLocationWiseStatus.QtyRcvd,VMaterialLocationWiseStatus.QtyIssued,VMaterialLocationWiseStatus.QtyPlaced,
						 VMaterialLocationWiseStatus.StoreName,VMaterialLocationWiseStatus.RackNo,VMaterialLocationWiseStatus.ShelfNo

FROM            dbo.ItemsRMComp INNER JOIN
                         dbo.VRM ON dbo.ItemsRMComp.RMID = dbo.VRM.RMID INNER JOIN
                         dbo.Items ON dbo.ItemsRMComp.ItemID = dbo.Items.ItemID LEFT OUTER JOIN
                         dbo.Processes ON dbo.ItemsRMComp.ProcessID = dbo.Processes.ProcessID
						 LEFT OUTER JOIN VMaterialLocationWiseStatus ON VRM.RMID1=VMaterialLocationWiseStatus.MaterialID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItemsWithLotWiseStock]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItemsWithLotWiseStock];
END
GO


CREATE VIEW [dbo].[VItemsWithLotWiseStock]
AS
SELECT dbo.VStoreShelfs.StoreName, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, T1.ItemID, T1.LotNo, T1.RcvdQty - ISNULL(T1.IssdQty, 0) AS NetQty, T1.Shelf_RefID, T1.Remarks, dbo.VStoreShelfs.Store_RefID, T1.EntryID
    ,CASE WHEN ISNULL(Lots_List.Batch_No,'')='' THEN 'No-Batch' ELSE ISNULL(Lots_List.Batch_No,'') END  AS Batch_No,ISNULL(Lots_List.Mill_Certificate_No,'') AS Mill_Certificate_No
FROM  dbo.VStoreShelfs INNER JOIN
             (SELECT dbo.RcvItemsSimpleDetail.EntryID, dbo.RcvItemsSimpleDetail.ItemID, ISNULL(dbo.vendRcvdDetail.LotNo, dbo.RcvItemsSimpleDetail.LotNo_Manual) AS LotNo, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID, dbo.RcvItemsSimpleDetail_Placement.RcvdQty, ISNULL(T1_1.IssdQty, 0) AS IssdQty, dbo.RcvItemsSimpleDetail_Placement.RcvdQty - ISNULL(T1_1.IssdQty, 0) 
                     AS NetQty, dbo.RcvItemsSimpleDetail.Remarks
            FROM  dbo.RcvItemsSimpleDetail INNER JOIN
                     dbo.RcvItemsSimpleDetail_Placement ON dbo.RcvItemsSimpleDetail.EntryID = dbo.RcvItemsSimpleDetail_Placement.RISD_RefID LEFT OUTER JOIN
                     dbo.TransferredToReadyFinishLotsDetail ON dbo.RcvItemsSimpleDetail.TTRFLD_RefID = dbo.TransferredToReadyFinishLotsDetail.EntryID LEFT OUTER JOIN
                     dbo.vendRcvdDetail ON dbo.TransferredToReadyFinishLotsDetail.VRD_RefID = dbo.vendRcvdDetail.EntryID LEFT OUTER JOIN
                         (SELECT RCV_ISD_RefID, SUM(Qty) AS IssdQty
                        FROM  dbo.IssItemsSimpleDetail
                        GROUP BY RCV_ISD_RefID) AS T1_1 ON dbo.RcvItemsSimpleDetail.EntryID = T1_1.RCV_ISD_RefID) AS T1 ON T1.Shelf_RefID = dbo.VStoreShelfs.EntryID
                        LEFT OUTER JOIN Lots_List ON T1.LotNo=Lots_List.LotNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VItemsWithShelfWiseStock_Old]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VItemsWithShelfWiseStock_Old];
END
GO

CREATE   VIEW [dbo].[VItemsWithShelfWiseStock_Old]
AS
SELECT dbo.VStoreShelfs.StoreName, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, T1.ItemID, T1.RcvdQty - ISNULL(T1.IssdQty, 0) AS NetQty, T1.Shelf_RefID, T2.Remarks, dbo.VStoreShelfs.Store_RefID
FROM  dbo.VStoreShelfs INNER JOIN
             (SELECT ItemID, Shelf_RefID, SUM(RcvdQty) AS RcvdQty, SUM(IssdQty) AS IssdQty
            FROM  (SELECT dbo.RcvItemsSimpleDetail.ItemID, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID, SUM(dbo.RcvItemsSimpleDetail_Placement.RcvdQty) AS RcvdQty, 0 AS IssdQty
                     FROM  dbo.RcvItemsSimpleDetail_Placement INNER JOIN
                              dbo.RcvItemsSimpleDetail ON dbo.RcvItemsSimpleDetail_Placement.RISD_RefID = dbo.RcvItemsSimpleDetail.EntryID
                     GROUP BY dbo.RcvItemsSimpleDetail.ItemID, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID
                     UNION ALL
                     SELECT dbo.IssItemsSimpleDetail.ItemID, dbo.IssItemsSimpleDetail_More.Shelf_RefID, 0 AS RcvdQty, SUM(dbo.IssItemsSimpleDetail_More.IssdQty) AS IssdQty
                     FROM  dbo.IssItemsSimpleDetail INNER JOIN
                              dbo.IssItemsSimpleDetail_More ON dbo.IssItemsSimpleDetail.EntryID = dbo.IssItemsSimpleDetail_More.IISD_RefID
                     GROUP BY dbo.IssItemsSimpleDetail.ItemID, dbo.IssItemsSimpleDetail_More.Shelf_RefID) AS TInner1
            GROUP BY ItemID, Shelf_RefID) AS T1 ON dbo.VStoreShelfs.EntryID = T1.Shelf_RefID LEFT OUTER JOIN
             (SELECT dbo.VRcvitemsSimpleDetail_Placement.ItemID, dbo.VRcvitemsSimpleDetail_Placement.Shelf_RefID, dbo.VRcvitemsSimpleDetail_Placement.Remarks
            FROM  dbo.VRcvitemsSimpleDetail_Placement INNER JOIN
                         (SELECT MAX(EntryID) AS MaxEntryID
                        FROM  dbo.VRcvitemsSimpleDetail_Placement AS VRcvitemsSimpleDetail_Placement_1
                        GROUP BY ItemID, Shelf_RefID) AS TMax ON dbo.VRcvitemsSimpleDetail_Placement.EntryID = TMax.MaxEntryID) AS T2 ON T1.ItemID = T2.ItemID AND T1.Shelf_RefID = T2.Shelf_RefID
WHERE (T1.RcvdQty - ISNULL(T1.IssdQty, 0) > 0)

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VLedger]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VLedger];
END
GO

-- ====================================================================================================
-- [15/49] SCRIPT: 19-07-26-11-VLedger.sql
-- ====================================================================================================
CREATE   VIEW [dbo].[VLedger] AS
WITH ChequeMapping AS (
    -- This isolates the VchrNo and finds the single populated CSNo, ignoring the NULLs
    SELECT 
        VchrNo, 
        MAX(CSNo) AS ValidCSNo
    FROM dbo.Vouchers
    WHERE CSNo IS NOT NULL
    GROUP BY VchrNo
)
SELECT 
    v.SNo, 
    v.VDate, 
    v.VchrNo, 
    v.Accno, 
    a.AccTitle, 
    v.Description, 
    v.Debit, 
    v.Credit, 
    v.balance, 
    v.CSNo, 
    v.DpstSlip, 
    vi.UserName, 
    vi.MachineName, 
    vi.DT, 
    c.chequeno, 
    c.chequeDate, 
    a.SubAccOf, 
    c.chequeType, 
    vi.Handed_Over_To, 
    c.Payee
FROM dbo.Vouchers v
INNER JOIN dbo.Accounts a 
    ON v.Accno = a.AccNo
-- Join to our new map to get the correct CSNo for every row of the voucher
LEFT OUTER JOIN ChequeMapping cm 
    ON v.VchrNo = cm.VchrNo
-- Join to the Cheque table using the mapped CSNo, not the row-level CSNo
LEFT OUTER JOIN dbo.Cheque c 
    ON cm.ValidCSNo = c.SNo
LEFT OUTER JOIN dbo.VoucherInfo vi 
    ON v.VchrNo = vi.VchrNo
WHERE v.SNo NOT IN (
    SELECT Sno 
    FROM dbo.ADVSno 
    WHERE AccNo = v.Accno
);


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VLedger_Original]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VLedger_Original];
END
GO

CREATE   VIEW [dbo].[VLedger_Original]
AS
SELECT     dbo.Vouchers.SNo, dbo.Vouchers.VDate, dbo.Vouchers.VchrNo, dbo.Vouchers.Accno, dbo.Accounts.AccTitle, dbo.Vouchers.Description, dbo.Vouchers.Debit, dbo.Vouchers.Credit, dbo.Vouchers.balance, dbo.Vouchers.CSNo, dbo.Vouchers.DpstSlip, 
                  dbo.VoucherInfo.UserName, dbo.VoucherInfo.MachineName, dbo.VoucherInfo.DT, dbo.Cheque.chequeno, dbo.Cheque.chequeDate, dbo.Accounts.SubAccOf, dbo.Cheque.chequeType, dbo.VoucherInfo.Handed_Over_To, dbo.Cheque.Payee
FROM        dbo.Accounts INNER JOIN
                  dbo.Vouchers ON dbo.Accounts.AccNo = dbo.Vouchers.Accno LEFT OUTER JOIN
                  dbo.Cheque ON dbo.Vouchers.CSNo = dbo.Cheque.SNo LEFT OUTER JOIN
                  dbo.VoucherInfo ON dbo.Vouchers.VchrNo = dbo.VoucherInfo.VchrNo
WHERE     (dbo.Vouchers.SNo NOT IN
                      (SELECT     Sno
                       FROM        dbo.ADVSno
                       WHERE     (AccNo = dbo.Vouchers.Accno)))

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VLot_With_ForgeBatchNo]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VLot_With_ForgeBatchNo];
END
GO

CREATE   VIEW [dbo].[VLot_With_ForgeBatchNo]
AS
SELECT dbo.vendRcvdDetail.LotNo, dbo.VendRcvdDetailPO.RcvID
FROM  dbo.VendIssdDetail INNER JOIN
         dbo.VendIssued ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
         dbo.RawMaterialIssuance ON dbo.VendIssued.EntryID = dbo.RawMaterialIssuance.VI_RefID INNER JOIN
         dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
         dbo.RMID_MLS_Details ON dbo.RawMaterialIssuanceDetail.EntryID = dbo.RMID_MLS_Details.RMID_RefID INNER JOIN
         dbo.MaterialLocationwiseStatus ON dbo.RMID_MLS_Details.MLS_RefID = dbo.MaterialLocationwiseStatus.EntryID INNER JOIN
         dbo.VendRcvdDetailPO ON dbo.MaterialLocationwiseStatus.Rcvd_RefID = dbo.VendRcvdDetailPO.EntryID INNER JOIN
         dbo.vendRcvdDetail ON dbo.VendIssdDetail.EntryID = dbo.vendRcvdDetail.Issue_RefID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VLot_With_TrayNo]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VLot_With_TrayNo];
END
GO

CREATE   VIEW [dbo].[VLot_With_TrayNo]
AS
SELECT dbo.vendRcvdDetail.LotNo, dbo.VendRcvdDetail_QC_Report.TrayNo
FROM  dbo.VendRcvdDetail_QC_Report INNER JOIN
         dbo.vendRcvdDetail ON dbo.VendRcvdDetail_QC_Report.VRD_RefID = dbo.vendRcvdDetail.EntryID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VLots_With_SOADetails]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VLots_With_SOADetails];
END
GO

CREATE   VIEW [dbo].[VLots_With_SOADetails]
AS
SELECT        TLots.LotNo, dbo.RcvItemsSimpleDetail.LotNo_Manual AS LotNo_SOA, dbo.Lots_List.Batch_No AS BatchNo_SOA, dbo.Lots_List.Mill_Certificate_No AS Mil_Certificate_No_SOA, 
                         dbo.StockOrderAdjustmentsDetail.Qty AS Qty_SOA
FROM            dbo.StockOrderAdjustments INNER JOIN
                             (SELECT        LotNo, MIN(EntryID) AS MinEntryID
                               FROM            dbo.VendRcvdDetail
                               GROUP BY LotNo) AS TLots ON dbo.StockOrderAdjustments.VID_RefID = TLots.MinEntryID INNER JOIN
                         dbo.StockOrderAdjustmentsDetail ON dbo.StockOrderAdjustments.EntryID = dbo.StockOrderAdjustmentsDetail.SOA_RefID INNER JOIN
                         dbo.RcvItemsSimpleDetail ON dbo.StockOrderAdjustmentsDetail.RCV_ISD_RefID = dbo.RcvItemsSimpleDetail.EntryID INNER JOIN
                         dbo.Lots_List ON dbo.RcvItemsSimpleDetail.LotNo_Manual = dbo.Lots_List.LotNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VMakerAdvancesLedger]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VMakerAdvancesLedger];
END
GO


CREATE VIEW [dbo].[VMakerAdvancesLedger]
AS
SELECT VendID,T1.* FROM
(SELECT EntryID,AccNo,DT,Description,Amount,AccVoucherNo,DAmount,'' AS Remarks FROM MakerAdvances
UNION
SELECT MakerAmtCleared.EntryID,MakerAmtCleared.AccNo,MakerAmtCleared.DT,'Deduction in Bill',-AmtClrd,MakerAmtCleared.VchrNo,0 AS DAmount,MakerLoanClearance_Manual.Remarks FROM MakerAmtCleared
LEFT JOIN MakerLoanClearance_Manual ON MakerAmtCleared.EntryID=MakerLoanClearance_Manual.MPB_STD_EntryID
) T1 INNER JOIN Makers ON T1.AccNo=Makers.AccNo




GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VMakerBalance_Long]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VMakerBalance_Long];
END
GO

CREATE   VIEW dbo.VMakerBalance_Long
AS
SELECT     AccNo, SUM(Amount) AS Balance
FROM        dbo.VMakerAdvancesLedger
GROUP BY AccNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VMakerBilling_Receivings_Hidden]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VMakerBilling_Receivings_Hidden];
END
GO


CREATE VIEW [dbo].[VMakerBilling_Receivings_Hidden]
AS
SELECT        dbo.MakerBilling_Receivings_Hidden.EntryID, dbo.MakerBilling_Receivings_Hidden.VRD_RefID, dbo.VVendReceivingList_Simple.DT, dbo.VVendReceivingList_Simple.ProcessID, dbo.VVendReceivingList_Simple.ItemCode, 
                         dbo.VVendReceivingList_Simple.RcvdQty, dbo.VVendReceivingList_Simple.VenderName, dbo.VVendReceivingList_Simple.VendID1, dbo.VVendReceivingList_Simple.OnlyDT, dbo.VVendReceivingList_Simple.Description, 
                         dbo.VVendReceivingList_Simple.LotNo, dbo.VVendReceivingList_Simple.OrderNo, dbo.VVendReceivingList_Simple.CompItemCode, dbo.VVendReceivingList_Simple.ReWorkLot, dbo.VVendReceivingList_Simple.ItemName, 
                         dbo.VVendReceivingList_Simple.ItemSize, dbo.VVendReceivingList_Simple.SizeUnit, dbo.VVendReceivingList_Simple.Type, dbo.VVendReceivingList_Simple.CustCode, dbo.VVendReceivingList_Simple.VendID
FROM            dbo.MakerBilling_Receivings_Hidden INNER JOIN
                         dbo.VVendReceivingList_Simple ON dbo.MakerBilling_Receivings_Hidden.VRD_RefID = dbo.VVendReceivingList_Simple.VRD_EntryID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VMakers]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VMakers];
END
GO

CREATE   VIEW [dbo].[VMakers]
AS
SELECT        dbo.Makers.VendID, dbo.Makers.VendID1, CASE WHEN dbo.Accounts.AccTitle IS NULL THEN VenderName ELSE Accounts.AccTitle END AS VenderName, dbo.Makers.Phone1, dbo.Makers.Phone2, dbo.Makers.Fax1, 
                         dbo.Makers.Address, dbo.Makers.ContactPerson, dbo.Makers.CPhone, dbo.Makers.CEmail, dbo.Makers.Mobile, dbo.Makers.NICNo, dbo.Makers.NTNNo, dbo.Makers.RefBy, dbo.Makers.RefByFName, dbo.Makers.RefByPhone1, 
                         dbo.Makers.RefByPhone2, dbo.Makers.RefByAddress, dbo.Makers.TimeLimit, dbo.Makers.MaxLimit, dbo.Makers.AccNo, dbo.Makers.VendType, dbo.Makers.PhaseID, dbo.Makers.SubVendType, dbo.Makers.Planter, 
                         dbo.Makers.Snaffer, dbo.Makers.Stamper, dbo.Makers.Experience, dbo.Makers.OtherCompany1, dbo.Makers.OtherCompany2, dbo.Makers.OpenBal, dbo.Makers.Active, dbo.Makers.OwnRepair, dbo.Makers.RepairVend, 
                         dbo.Makers.RepairDedRate, dbo.Makers.MakerNameUrdu, dbo.Makers.BankAccNo, dbo.Accounts.AccTitle, dbo.Processes.Description, dbo.Makers.VenderName AS MakerName, dbo.Makers.MakerType, dbo.Accounts.Balance, 
                         dbo.Makers.VendPic, dbo.Makers.VendThumb, dbo.Makers.AuthRequired, dbo.Makers.ExcessQtyPercentage, dbo.Makers.MaximumRcvingsAgainstPO, dbo.Makers.ShowRateOnPO, dbo.Makers.PaymentTerms, 
                         dbo.Accounts.SubAccOf, dbo.Makers.MakerCapacity, dbo.Makers.CompanyName, dbo.Makers.Maker_Second_Name
						 ,dbo.Makers.CNIC_PDF_FileName
FROM            dbo.Makers LEFT OUTER JOIN
                         dbo.Accounts ON dbo.Makers.AccNo = dbo.Accounts.AccNo LEFT OUTER JOIN
                         dbo.Processes ON dbo.Makers.PhaseID = dbo.Processes.ProcessID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VMaterialLocationWiseStatus]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VMaterialLocationWiseStatus];
END
GO


CREATE VIEW [dbo].[VMaterialLocationWiseStatus]
AS
SELECT     dbo.MaterialLocationwiseStatus.EntryID, dbo.MaterialLocationwiseStatus.Rcvd_RefID, dbo.MaterialLocationwiseStatus.Shelf_RefID, dbo.MaterialLocationwiseStatus.QtyPlaced, dbo.MaterialLocationwiseStatus.QtyIssued, dbo.VendRcvdDetailPO.MaterialID, 
                  dbo.MaterialLocationwiseStatus.SheetsPlaced, dbo.MaterialLocationwiseStatus.SheetsIssued, dbo.RM.RMName, dbo.VVenders.AccTitle, dbo.VendRcvdDetailPO.QtyRcvd, dbo.MaterialLocationwiseStatus.BatchNo, dbo.MaterialLocationwiseStatus.LotNo, 
                  dbo.VStoreShelfs.StoreName, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, dbo.VendRcvdDetailPO.RcvID, dbo.VendRcvdDetailPO.PORefNo, dbo.VendRcvd.RcvDate, dbo.MaterialLocationwiseStatus.Mill_Certificate_No, dbo.VVenders.AccNo, dbo.RM.Pic, 
                  dbo.RM.Unit, dbo.RM.GroupID, dbo.RM.RMID1, dbo.RM.TechnicalDrawing, dbo.MaterialLocationwiseStatus.UserName, dbo.MaterialLocationwiseStatus.MachineName, dbo.VVenders.VendID, dbo.VVenders.VenderDescription, dbo.VVenders.ContactPerson, 
                  dbo.VVenders.CPhone
FROM        dbo.MaterialLocationwiseStatus INNER JOIN
                  dbo.VendRcvdDetailPO ON dbo.MaterialLocationwiseStatus.Rcvd_RefID = dbo.VendRcvdDetailPO.EntryID INNER JOIN
                  dbo.RM ON dbo.VendRcvdDetailPO.MaterialID = dbo.RM.RMID1 INNER JOIN
                  dbo.VendRcvd ON dbo.VendRcvdDetailPO.RcvID = dbo.VendRcvd.RcvID INNER JOIN
                  dbo.VStoreShelfs ON dbo.MaterialLocationwiseStatus.Shelf_RefID = dbo.VStoreShelfs.EntryID LEFT OUTER JOIN
                  dbo.VVenders ON dbo.VendRcvd.VendID = dbo.VVenders.AccNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER OFF;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VMiscEntries_GateInTemp]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VMiscEntries_GateInTemp];
END
GO

CREATE   VIEW [dbo].[VMiscEntries_GateInTemp] AS 
SELECT m.EntryID, m.GatePassNo, m.DT, m.PersonName, m.VehicleNo, m.Description, m.UserName, m.MachineName, 
m.DTEntry, CAST(CONVERT(VARCHAR(50), m.DT, 6) AS DATETIME) AS OnlyDT, m.EmpID, emp.Name AS EmpName, m.GatePassType, m.GatePassType2, 
m.DeptID, d.name AS DeptName, m.VendID, vm.AccTitle AS Name, m.VendNameManual, m.VM_AccNo, ISNULL(m.Approved_HOD, 0) AS Approved_HOD
, ISNULL(m.Approved, 0) AS Approved, ISNULL(m.Approved_Finance, 0) AS Approved_Finance FROM dbo.MiscEntries_GateInTemp m 
LEFT OUTER JOIN dbo.Departments d ON m.DeptID = d.deptid 
LEFT OUTER JOIN dbo.Employees emp ON m.EmpID = emp.EmpID 
LEFT OUTER JOIN dbo.VVendersAndMakers vm ON m.VM_AccNo = vm.AccNo;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VMiscEntries_GateOut]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VMiscEntries_GateOut];
END
GO

CREATE   VIEW dbo.VMiscEntries_GateOut
AS
SELECT 
    m.EntryID, 
    m.GatePassNo, 
    m.DT, 
    m.PersonName, 
    m.VehicleNo, 
    m.Description, 
    m.UserName, 
    m.MachineName, 
    m.DTEntry, 
    CAST(CONVERT(VARCHAR(50), m.DT, 6) AS DATETIME) AS OnlyDT, 
    m.EmpID, 
    m.GatePassType, 
    m.DeptID, 
    m.VendID, 
    vm.AccTitle AS Name, 
    m.VendNameManual,
    m.VM_AccNo,
    ISNULL(m.Approved_HOD, 0) AS Approved_HOD,
    ISNULL(m.Approved, 0) AS Approved,
    ISNULL(m.Approved_Finance, 0) AS Approved_Finance
FROM dbo.MiscEntries_GateOut m
LEFT OUTER JOIN dbo.VVendersAndMakers vm ON m.VM_AccNo = vm.AccNo;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VMonthlySalaries]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VMonthlySalaries];
END
GO

CREATE   VIEW [dbo].[VMonthlySalaries]
AS
SELECT        dbo.MonthlySalaries.EntryID, dbo.MonthlySalaries.EmpID, dbo.MonthlySalaries.DT, dbo.MonthlySalaries.IsForSA, dbo.MonthlySalaries.BSal, dbo.MonthlySalaries.Rate, dbo.MonthlySalaries.ADays, dbo.MonthlySalaries.AAmt, 
                         dbo.MonthlySalaries.AAllow, dbo.MonthlySalaries.AAllowAmt, dbo.MonthlySalaries.SDays, dbo.MonthlySalaries.SAmt, dbo.MonthlySalaries.Leaves, dbo.MonthlySalaries.LeaveAmt, dbo.MonthlySalaries.OHrs, 
                         dbo.MonthlySalaries.OAmt, dbo.MonthlySalaries.LHrs, dbo.MonthlySalaries.LAmt, dbo.MonthlySalaries.Total, dbo.MonthlySalaries.Tax, dbo.MonthlySalaries.NetTtl, dbo.MonthlySalaries.ShortTerm, 
                         dbo.MonthlySalaries.LongTerm, dbo.MonthlySalaries.AdvSal, dbo.MonthlySalaries.Unionfund, dbo.MonthlySalaries.Fine, dbo.MonthlySalaries.Bonus, dbo.MonthlySalaries.Lunch, dbo.MonthlySalaries.EOBI, 
                         dbo.MonthlySalaries.Balance, dbo.MonthlySalaries.PrevLTLoan, dbo.MonthlySalaries.CasualLeaves, dbo.MonthlySalaries.SickLeaves, dbo.MonthlySalaries.AnnualLeaves, dbo.MonthlySalaries.CompensatoryLeaves, 
                         dbo.MonthlySalaries.WPLeaves, dbo.MonthlySalaries.MaternityLeaves, dbo.MonthlySalaries.HrsPerDay, dbo.MonthlySalaries.TotalMonthHrs, dbo.MonthlySalaries.DTFinal, dbo.MonthlySalaries.FakeWorkingHrs, 
                         dbo.MonthlySalaries.FakeRate, dbo.MonthlySalaries.FakeSalary, dbo.MonthlySalaries.SundayOTHrs, dbo.MonthlySalaries.SundayOTRate, dbo.MonthlySalaries.FixAllowance, dbo.MonthlySalaries.HoldSalaryAmt, 
                         dbo.MonthlySalaries.PresentDays, dbo.MonthlySalaries.LeaveDays, dbo.MonthlySalaries.LateComingHrs, dbo.MonthlySalaries.ShortHrs, dbo.MonthlySalaries.AmtPaid, dbo.MonthlySalaries.GPHrs, 
                         dbo.MonthlySalaries.GPHrsAmt, dbo.MonthlySalaries.DeptID, dbo.MonthlySalaries.OTDinnerCount, dbo.MonthlySalaries.OTDinnerAmount, dbo.MonthlySalaries.DedOnePercent, dbo.MonthlySalaries.PerformanceDedAmt, 
                         dbo.MonthlySalaries.RejectionDedAmt, dbo.VEmp.name, dbo.VEmp.fname, dbo.VEmp.EmpType, dbo.VEmp.EmpAccNo, dbo.VEmp.BankPymt, dbo.VEmp.Rel, dbo.Departments.name AS DeptName, 
                         dbo.MonthlySalaries.ShortHrsAmt, dbo.MonthlySalaries.ZeroAbsentBonus, dbo.MonthlySalaries.OTHrs_Original, dbo.MonthlySalaries.LateHrs_Original, dbo.MonthlySalaries.OTHrs_Net, dbo.MonthlySalaries.LateHrs_Net
						 ,dbo.SalaryAlreadyPaid.AmtPaid AS Salary_Already_Paid_Amt
						 ,dbo.MonthlySalaries.Salary_Paid,dbo.MonthlySalaries.Salary_Paid_UserName,dbo.MonthlySalaries.Salary_Paid_MachineName,dbo.MonthlySalaries.Salary_Paid_DTEntry
FROM            dbo.MonthlySalaries INNER JOIN
                         dbo.VEmp ON dbo.MonthlySalaries.EmpID = dbo.VEmp.empid INNER JOIN
                         dbo.Departments ON dbo.MonthlySalaries.DeptID = dbo.Departments.deptid
						 LEFT JOIN dbo.SalaryAlreadyPaid ON dbo.MonthlySalaries.EmpID=dbo.SalaryAlreadyPaid.EmpID AND dbo.MonthlySalaries.DT=dbo.SalaryAlreadyPaid.ForMonth

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VPaySlip]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VPaySlip];
END
GO


CREATE VIEW [dbo].[VPaySlip]
AS
SELECT        dbo.Departments.deptid, dbo.Departments.name AS DeptName, dbo.Employees.empid, dbo.Employees.name, dbo.Employees.Rel, dbo.Employees.fname, dbo.PrintSalary.BSal, dbo.PrintSalary.Rate, dbo.PrintSalary.ADays, 
                         dbo.PrintSalary.AAmt, dbo.PrintSalary.SDays, dbo.PrintSalary.SAmt, dbo.PrintSalary.OHrs, dbo.PrintSalary.OAmt, dbo.PrintSalary.LHrs, dbo.PrintSalary.LAmt, dbo.PrintSalary.Total, dbo.PrintSalary.Tax, dbo.PrintSalary.NetTtl, 
                         dbo.PrintSalary.Paid, dbo.PrintSalary.Balance, dbo.PrintSalary.DT, dbo.Employees.Salary, dbo.Employees.PProfit, dbo.Employees.StartingSalary, dbo.PrintSalary.LongTerm, dbo.Employees.Designation, 
                         dbo.PrintSalary.AdvSal, dbo.Employees.EmpType, dbo.Employees.Status, dbo.PrintSalary.PrevLTLoan, dbo.PrintSalary.AdvPer, dbo.PrintSalary.AAllow, dbo.PrintSalary.AAllowAmt, dbo.PrintSalary.Leaves, 
                         dbo.PrintSalary.LeaveAmt, dbo.PrintSalary.Fine, dbo.PrintSalary.Bonus, dbo.PrintSalary.SocialAmt, dbo.Employees.NTN, dbo.Employees.BankPymt, dbo.PrintSalary.CasualLeaves, dbo.PrintSalary.SickLeaves, 
                         dbo.PrintSalary.AnnualLeaves, dbo.PrintSalary.CompensatoryLeaves, dbo.PrintSalary.MaternityLeaves, dbo.PrintSalary.WPLeaves, dbo.Employees.sex, dbo.Employees.Lunch, dbo.PrintSalary.EOBI, dbo.Employees.EmpIDOld, 
                         dbo.PrintSalary.SundayOTHrs, dbo.PrintSalary.SundayOTRate, dbo.PrintSalary.FixAllowance, dbo.Employees.EOBIAmt, dbo.Employees.SocialSecurityAmt, dbo.Employees.EmpLunchAmt, dbo.PrintSalary.HoldSalaryAmt, 
                         dbo.PrintSalary.PresentDays, dbo.PrintSalary.LeaveDays, dbo.PrintSalary.LateComingHrs, dbo.PrintSalary.ShortHrs, dbo.PrintSalary.AmtPaid, dbo.PrintSalary.GPHrs, dbo.PrintSalary.GPHrsAmt, dbo.PrintSalary.OTDinnerCount, 
                         dbo.PrintSalary.OTDinnerAmount, dbo.PrintSalary.DedOnePercent, dbo.PrintSalary.PerformanceDedAmt, dbo.PrintSalary.RejectionDedAmt, dbo.PrintSalary.ShortHrsAmt, dbo.PrintSalary.ZeroAbsentBonus, 
                         dbo.VEmpSettings.OTRate, dbo.PrintSalary.OTHrs_Original, dbo.PrintSalary.LateHrs_Original, dbo.PrintSalary.OTHrs_Net, dbo.PrintSalary.LateHrs_Net
FROM            dbo.Departments INNER JOIN
                         dbo.PrintSalary INNER JOIN
                         dbo.Employees ON dbo.PrintSalary.EmpID = dbo.Employees.empid ON dbo.Departments.deptid = dbo.PrintSalary.DeptID INNER JOIN
                         dbo.VEmpSettings ON dbo.Employees.empid = dbo.VEmpSettings.empid

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VProcesses]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VProcesses];
END
GO

    CREATE   VIEW dbo.VProcesses
    AS
    SELECT     dbo.Processes.ProcessID, dbo.Processes.SNO, dbo.Processes.Description, dbo.Processes.Supervisor, dbo.Processes.Operation, dbo.Processes.AuthRequired, 
                          dbo.Processes.Code, dbo.Processes.ProcessNameUrdu, dbo.Processes.Insp_RefID, dbo.Processes.Fix_Maker_RefID, dbo.Makers.VenderName, 
                          dbo.InspectionProcesses.Code AS Insp_Code, dbo.InspectionProcesses.ProcessName AS Insp_ProcessName, 
                          dbo.InspectionProcesses.ProcessNameUrdu AS Insp_ProcessNameUrdu, dbo.Processes.InspectionProcess, dbo.Processes.ProcessNameUrduOther, 
                          dbo.Processes.BillingProcessID,
                          ISNULL(dbo.Processes.EstimatedMinutes, 0) AS EstimatedMinutes
    FROM         dbo.Processes LEFT OUTER JOIN
                          dbo.InspectionProcesses ON dbo.Processes.Insp_RefID = dbo.InspectionProcesses.EntryID LEFT OUTER JOIN
                          dbo.Makers ON dbo.Processes.Fix_Maker_RefID = dbo.Makers.VendID
    
GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRawMaterialIssuanceList]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRawMaterialIssuanceList];
END
GO

CREATE   VIEW [dbo].[VRawMaterialIssuanceList]
AS
SELECT        dbo.RawMaterialIssuance.IssNo, dbo.RawMaterialIssuance.DT, dbo.RawMaterialIssuance.EntryDT, dbo.RawMaterialIssuance.UserName, dbo.RawMaterialIssuance.ComputerName, dbo.RawMaterialIssuance.VI_RefID, 
                         dbo.RawMaterialIssuance.IssdToEmpID, dbo.RawMaterialIssuance.VchrNo, dbo.RawMaterialIssuance.VO_OrderNo, dbo.RawMaterialIssuance.AuthorizedByEmpID, dbo.RawMaterialIssuance.EntryID, 
                         dbo.VEmp.name AS IssdToEmpName, dbo.Makers.VenderName, TInner1.RM, dbo.VEmp.deptid, dbo.VendIssued.VendID, dbo.RawMaterialIssuance.Remarks, dbo.RawMaterialIssuance.Entry_Type, 
                         dbo.RawMaterialIssuance.Vendor_AccNo, dbo.RawMaterialIssuance.Vendor_PO_No, dbo.RawMaterialIssuance.Vendor_VchrNo,
                         VVendersAndMakers.AccTitle AS Debit_Note_Vendor,
                         TInner1.Total_Issued_Qty,
                         TEmpID.IssdToEmpIDOnly
FROM            dbo.Makers RIGHT OUTER JOIN
                         dbo.VendIssued ON dbo.Makers.VendID = dbo.VendIssued.VendID RIGHT OUTER JOIN
                         dbo.RawMaterialIssuance ON dbo.VendIssued.EntryID = dbo.RawMaterialIssuance.VI_RefID
                         OUTER APPLY (SELECT CASE WHEN CHARINDEX('}', dbo.RawMaterialIssuance.IssdToEmpID) > 2 
                                                  THEN SUBSTRING(dbo.RawMaterialIssuance.IssdToEmpID, 2, CHARINDEX('}', dbo.RawMaterialIssuance.IssdToEmpID) - 2)
                                                  ELSE dbo.RawMaterialIssuance.IssdToEmpID END AS IssdToEmpIDOnly) TEmpID
                         LEFT OUTER JOIN
                         dbo.VEmp ON TEmpID.IssdToEmpIDOnly = dbo.VEmp.empid LEFT OUTER JOIN
                             (SELECT        IssNo, MAX('{' + RMID1 + '} ' + RMName) AS RM, SUM(Qty) AS Total_Issued_Qty
                               FROM            dbo.VRawMaterialIssuance
                               GROUP BY IssNo) AS TInner1 ON dbo.RawMaterialIssuance.IssNo = TInner1.IssNo
LEFT JOIN VVendersAndMakers ON dbo.RawMaterialIssuance.Vendor_AccNo=VVendersAndMakers.AccNo
GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRcvitemsSimpleDetail_Placement]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRcvitemsSimpleDetail_Placement];
END
GO


CREATE VIEW [dbo].[VRcvitemsSimpleDetail_Placement]
AS
SELECT        dbo.RcvItemsSimpleDetail.ItemID, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID, dbo.RcvItemsSimpleDetail_Placement.RcvdQty, dbo.RcvItemsSimpleDetail_Placement.IssdQty, 
                         dbo.RcvItemsSimpleDetail_Placement.EntryID, dbo.RcvItemsSimpleDetail_Placement.RISD_RefID, dbo.RcvItemsSimpleDetail_Placement.Remarks, dbo.RcvItemsSimpleDetail.Sample_Type, 
                         dbo.RcvItemsSimple.CustCode, dbo.VStoreShelfs.StoreName, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, dbo.RcvItemsSimpleDetail.TTRFLD_RefID
FROM            dbo.RcvItemsSimpleDetail INNER JOIN
                         dbo.RcvItemsSimpleDetail_Placement ON dbo.RcvItemsSimpleDetail.EntryID = dbo.RcvItemsSimpleDetail_Placement.RISD_RefID INNER JOIN
                         dbo.RcvItemsSimple ON dbo.RcvItemsSimpleDetail.RcvNo = dbo.RcvItemsSimple.RcvNo LEFT OUTER JOIN
         dbo.VStoreShelfs ON dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID = dbo.VStoreShelfs.EntryID


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRM]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRM];
END
GO
-- ====================================================================================================
-- PART 3: PREREQUISITE SYNCHRONIZED VIEWS
-- ====================================================================================================

CREATE   VIEW [dbo].[VRM]
AS
SELECT        ISNULL(dbo.VRMAvgRate.AvgRate, 0) AS AvgRate, dbo.RMGroups.Description AS GroupName, dbo.RMGroups.GrpColor AS GroupColor, dbo.RM.RMID, dbo.RM.GroupID, dbo.RM.RMID1, dbo.RM.RMName, dbo.RM.Unit, 
                         dbo.RM.RMSize, dbo.RM.SizeUnit, dbo.RM.Rate, dbo.RM.RMUsage, dbo.RM.AssetAccNo, dbo.RM.SaleAccNo, dbo.RM.SRTAccNo, dbo.RM.InActive, dbo.RM.ReorderPoint, dbo.RM.MinLevel, dbo.RM.MaxLevel, 
                         dbo.RM.OpeningStock, dbo.RM.Description, dbo.RM.MaterialCode, dbo.RM.TechnicalDrawing, dbo.RM.Pic, dbo.RM.AnnealingStock, dbo.RM.MachiningStock, dbo.RM.MakerRate, dbo.RM.UrduName, Tab1.QtyInStock, 
                         dbo.RM.RequiresMaleGrinding, dbo.RM.MaleGrindingStock, dbo.RMGroupIDsForForging.Group_ID AS ForgingGroupID, Tab1.SheetsInStock, dbo.RM.Returnable, dbo.RM.Sampling, dbo.RM.RM_Maker_Description, 
                         dbo.RM.Weight_Length
FROM            dbo.RM LEFT OUTER JOIN
                         dbo.RMGroupIDsForForging ON dbo.RM.GroupID = dbo.RMGroupIDsForForging.Group_ID LEFT OUTER JOIN
                         dbo.RMGroups ON dbo.RM.GroupID = dbo.RMGroups.ID LEFT OUTER JOIN
                         dbo.VRMAvgRate ON dbo.RM.RMID = dbo.VRMAvgRate.RMID LEFT OUTER JOIN
                             (SELECT        MaterialID, SUM(QtyPlaced - QtyIssued) AS QtyInStock, SUM(ISNULL(SheetsPlaced, 0) - ISNULL(SheetsIssued, 0)) AS SheetsInStock
                               FROM            dbo.VMaterialLocationWiseStatus
                               GROUP BY MaterialID) AS Tab1 ON dbo.RM.RMID1 = Tab1.MaterialID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRMWithRMOpenPOs]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRMWithRMOpenPOs];
END
GO

CREATE   VIEW [dbo].[VRMWithRMOpenPOs]
AS
SELECT        dbo.VRM.AvgRate, dbo.VRM.GroupName, dbo.VRM.GroupColor, dbo.VRM.RMID, dbo.VRM.GroupID, dbo.VRM.RMID1, dbo.VRM.RMName, dbo.VRM.RMSize, dbo.VRM.SizeUnit, dbo.VRM.Unit, dbo.VRM.Rate, 
                         dbo.VRM.RMUsage, dbo.VRM.AssetAccNo, dbo.VRM.SaleAccNo, dbo.VRM.SRTAccNo, dbo.VRM.InActive, dbo.VRM.ReorderPoint, dbo.VRM.MinLevel, dbo.VRM.MaxLevel, dbo.VRM.OpeningStock, 
                         dbo.VRM.Description, dbo.VRM.MaterialCode, dbo.VRM.TechnicalDrawing, dbo.VRM.Pic, dbo.VRM.AnnealingStock, dbo.VRM.MakerRate, dbo.VRM.MachiningStock, dbo.VRM.UrduName, dbo.VRM.QtyInStock, 
                         dbo.VRM.RequiresMaleGrinding, dbo.VRM.MaleGrindingStock, dbo.VRM.ForgingGroupID, dbo.VRM.SheetsInStock, TOpenPos.OpenPOsQty, dbo.VRM.Sampling
						 ,dbo.VRM.Weight_Length
FROM            dbo.VRM LEFT OUTER JOIN
                             (SELECT        RMID1, SUM(QtyToRcv) AS OpenPOsQty
                               FROM            dbo.VVendOrdersToRcv
                               GROUP BY RMID1) AS TOpenPos ON dbo.VRM.RMID1 = TOpenPos.RMID1

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VrptCustomInvoiceDetail]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VrptCustomInvoiceDetail];
END
GO
CREATE   VIEW [dbo].[VrptCustomInvoiceDetail]
AS
SELECT dbo.FOrderItems.ID, dbo.FOrderItems.ItemCode, dbo.FCustomerCatalog.Description, dbo.FCustomerCatalog.ItemID, dbo.FCustomerCatalog.CompItemID, dbo.ForeignCustomers.Curr, dbo.FCustomerCatalog.Unit, dbo.ForeignCustomers.Name, dbo.ForeignCustomers.Address, dbo.FOrderItems.OrderNo, dbo.ForeignCustomers.Cont1name, dbo.ForeignCustomers.DTFormat, 
         dbo.FCustomerCatalog.PackingMode, dbo.VItems1.CompleteItemName, dbo.FProformaOrders.Qty AS ProformaQty, dbo.ForeignCustomers.FakeAddress, dbo.ItemCatagories.Description AS ItemCatagory, dbo.ItemCatagories.SRONo, dbo.FCustomerOrders.DeliveryDT, dbo.FProformaOrders.EntryID, dbo.FCustomerOrders.DT, dbo.VItems1.ItemGroup, 
         dbo.CustomInvoice.CustomInvoice, dbo.CustomInvoiceItems.EntryID AS CustomInvoiceItemsEntryID, dbo.CustomInvoiceItems.Qty, dbo.CustomInvoiceItems.Price, dbo.CustomInvoiceItems.CustomPrice, dbo.VItems1.CatID, dbo.VItems1.ItemName, dbo.VItems1.CustomDescription, dbo.VItems1.FinishedWeight, dbo.VItems1.ItemSize, dbo.VItems1.SizeUnit, 
         dbo.VItems1.Unit AS ItemUnit, dbo.VItems1.ItemColor, dbo.CustomInvoice.Consignee, dbo.FCustomerCatalog.ItemColor AS CustomerItemColor, dbo.VItems1.Type, dbo.CustomInvoiceItems.DTRENo, dbo.CustomInvoice.DTREDescription, dbo.CustomInvoiceItems.BatchNo, dbo.VItems1.FDAListingNo, dbo.VItems1.FDAProductCode, dbo.CustomInvoice.GrossWeight, 
         dbo.BatchNosFromPackingListForInvoice_F(dbo.CustomInvoiceItems.EntryID) AS Batches, dbo.CustomInvoice.ComDrawnUnder, dbo.CustomInvoice.ComSpecial, dbo.FCustomerOrders.InternalRefNo, dbo.FProformaOrders.ItemDescription, dbo.FOrderItems.Item_Finishing_Type
		 ,dbo.FCustomerCatalog.MDMA,dbo.FCustomerCatalog.SFDA_Listing_No,dbo.FCustomerCatalog.MD_Group,
		 LEFT(ItemCatagories.HSCode,4) AS HSCode,FOrderItems.IW_BatchNo,FOrderItems.IW_OrderNo,FOrderItems.Remarks,FOrderItems.Stamps,
		 FOrderItems.Quality,VItems1.SteelType,VItems1.SteelUsed,VItems1.SteelGage,CustomInvoice.DT AS InvoiceDT
FROM  dbo.VItems1 INNER JOIN
         dbo.CustomInvoiceItems INNER JOIN
         dbo.ForeignCustomers INNER JOIN
         dbo.FProformaOrders INNER JOIN
         dbo.FOrderItems ON dbo.FProformaOrders.OrderEntryID = dbo.FOrderItems.ID INNER JOIN
         dbo.FCustomerOrders ON dbo.FOrderItems.OrderNo = dbo.FCustomerOrders.OrderNo ON dbo.ForeignCustomers.CustCode = dbo.FCustomerOrders.CustCode AND dbo.ForeignCustomers.Country = dbo.FCustomerOrders.Country INNER JOIN
         dbo.CustomInvoice INNER JOIN
         dbo.FCustomerCatalog ON dbo.CustomInvoice.CustCode = dbo.FCustomerCatalog.CustCode AND dbo.CustomInvoice.Country = dbo.FCustomerCatalog.Country ON dbo.ForeignCustomers.CustCode = dbo.CustomInvoice.CustCode AND dbo.ForeignCustomers.Country = dbo.CustomInvoice.Country ON dbo.CustomInvoiceItems.RefID = dbo.FProformaOrders.EntryID AND 
         dbo.CustomInvoiceItems.CustomInvoice = dbo.CustomInvoice.CustomInvoice ON dbo.VItems1.ItemID = dbo.FOrderItems.CompItemCode AND dbo.VItems1.ItemID = dbo.FCustomerCatalog.CompItemID INNER JOIN
         dbo.ItemCatagories ON dbo.VItems1.CatID = dbo.ItemCatagories.CatID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VrptOrders]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VrptOrders];
END
GO


CREATE VIEW [dbo].[VrptOrders]
AS
SELECT dbo.FOrderItems.ID, dbo.FOrderItems.ItemCode, dbo.FOrderItems.Price, dbo.FOrderItems.Qty, dbo.FOrderItems.Qty - ISNULL(TProforma.ProformaUsedQty, 0) AS InvQty, dbo.Items.ItemName, dbo.FCustomerCatalog.Description, dbo.FCustomerCatalog.ItemID, dbo.FCustomerCatalog.CompItemID, dbo.ForeignCustomers.Curr, dbo.FCustomerCatalog.Unit, 
         dbo.ForeignCustomers.Name, dbo.ForeignCustomers.Address, dbo.FOrderItems.CustomPrice, dbo.Items.CustomDescription, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.SteelTypes.SteelType, dbo.SteelGages.Gage, dbo.Items.ItemColor, dbo.FCustomerCatalog.BarcodeNo, dbo.FCustomerCatalog.PackingMode, dbo.Items.MasterCartonL, dbo.Items.MasterCartonW, 
         dbo.Items.MasterCartonH, dbo.Items.SmallCartonL, dbo.Items.SmallCartonW, dbo.Items.SmallCartonH, dbo.Items.PolyBag, dbo.FCustomerCatalog.SpecialInstructions, dbo.FCustomerCatalog.StampInstructions, dbo.FCustomerCatalog.PackingInstructions, dbo.Items.Unit AS UnitItems, dbo.FCustomerCatalog.ItemColor AS CustomerColor, ISNULL(Shipped.TotalShipped, 0) 
         AS ShippedQty, dbo.Items.ItemNameUrdu, dbo.Items.InHand, dbo.FOrderItems.SortNo, dbo.ItemCatagories.MaxLotSize, dbo.FOrderItems.DeliveryDT AS DeliveryDTItem, dbo.FOrderItems.Stamps AS StampsItem, dbo.FOrderItems.Quality AS QualityItem, dbo.Items.FinQuality, dbo.Items.Type, dbo.FOrderItems.DeliveryStatus, dbo.FCustomerCatalog.OCR, 
         dbo.FOrderItems.CompItemCode, dbo.ItemGroups.Description AS GroupDescription, dbo.FCustomerOrders.OrderNo, dbo.FCustomerOrders.DT, dbo.FCustomerOrders.TradeTerms, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.PartialShipment, dbo.FCustomerOrders.PaymentTerms, dbo.FCustomerOrders.TransShipment, 
         dbo.FCustomerOrders.Packaging, dbo.FCustomerOrders.DeliveryDT, dbo.FCustomerOrders.BatchNo, dbo.FCustomerOrders.CompanyRefID, dbo.FCustomerOrders.StampDT, dbo.FCustomerOrders.Quality, dbo.FCustomerOrders.InternalRefNo, dbo.FCustomerOrders.OrderRcvdVia, dbo.FOrderItems.Remarks, dbo.FCustomerOrders.OrderType, dbo.Items.GroupID, T1.FOI_RefID, 
         dbo.FCustomerFinalOrders.Cancelled, dbo.Items.MasterCartonSmallBoxes, dbo.Items.SmallBoxPcs, dbo.Items.FinishedWeight, dbo.FCustomerOrders.OrderRevisionNo, dbo.FCustomerOrders.OrderRevisionDT, dbo.FOrderItems.Item_Edited, dbo.FCustomerOrders.ShippingMode, dbo.FOrderItems.Weight, dbo.FOrderItems.Item_Finishing_Type, 
         dbo.Item_Finishing_Type.Item_Finishing_Type_Text
         , dbo.FOrderItems.IW_OrderNo, dbo.FOrderItems.IW_BatchNo, dbo.Items.ItemPic
				  ,dbo.FCustomerOrders.Packaging_Weight,ReadyFinishPrice,Fillingprice,PriceForCost
FROM  dbo.FCustomerOrders INNER JOIN
         dbo.FOrderItems ON dbo.FCustomerOrders.OrderNo = dbo.FOrderItems.OrderNo INNER JOIN
         dbo.FCustomerCatalog ON dbo.FCustomerOrders.CustCode = dbo.FCustomerCatalog.CustCode AND dbo.FCustomerOrders.Country = dbo.FCustomerCatalog.Country AND dbo.FOrderItems.CompItemCode = dbo.FCustomerCatalog.CompItemID INNER JOIN
         dbo.Items ON dbo.FCustomerCatalog.CompItemID = dbo.Items.ItemID INNER JOIN
         dbo.ForeignCustomers ON dbo.FCustomerOrders.CustCode = dbo.ForeignCustomers.CustCode AND dbo.FCustomerOrders.Country = dbo.ForeignCustomers.Country INNER JOIN
         dbo.ItemCatagories ON dbo.Items.CatID = dbo.ItemCatagories.CatID INNER JOIN
         dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID LEFT OUTER JOIN
         dbo.Item_Finishing_Type ON dbo.FOrderItems.Item_Finishing_Type = dbo.Item_Finishing_Type.EntryID LEFT OUTER JOIN
         dbo.FCustomerFinalOrders ON dbo.FCustomerOrders.OrderNo = dbo.FCustomerFinalOrders.OrderNo LEFT OUTER JOIN
         dbo.SteelGages ON dbo.Items.Gage = dbo.SteelGages.GageID LEFT OUTER JOIN
         dbo.SteelTypes ON dbo.Items.SteelUsed = dbo.SteelTypes.SteelID LEFT OUTER JOIN
             (SELECT FOI_RefID
            FROM  dbo.FOrderItems_PDF
            GROUP BY FOI_RefID) AS T1 ON dbo.FOrderItems.ID = T1.FOI_RefID LEFT OUTER JOIN
             (SELECT OrderEntryID, SUM(Qty) AS ProformaUsedQty
            FROM  dbo.FProformaOrders
            GROUP BY OrderEntryID) AS TProforma ON dbo.FOrderItems.ID = TProforma.OrderEntryID LEFT OUTER JOIN
             (SELECT FProformaOrders_1.OrderEntryID, SUM(ISNULL(dbo.CustomInvoiceItems.Qty, 0)) AS TotalShipped
            FROM  dbo.CustomInvoiceItems INNER JOIN
                     dbo.CustomInvoice ON dbo.CustomInvoiceItems.CustomInvoice = dbo.CustomInvoice.CustomInvoice INNER JOIN
                     dbo.FProformaOrders AS FProformaOrders_1 ON dbo.CustomInvoiceItems.RefID = FProformaOrders_1.EntryID
            WHERE (dbo.CustomInvoice.GatePassDT IS NOT NULL)
            GROUP BY FProformaOrders_1.OrderEntryID) AS Shipped ON dbo.FOrderItems.ID = Shipped.OrderEntryID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VrptOrders_ForProduction_Simple]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VrptOrders_ForProduction_Simple];
END
GO

CREATE   VIEW [dbo].[VrptOrders_ForProduction_Simple]
AS
SELECT TFOrderItems.ItemCode, TFOrderItems.Qty, dbo.FCustomerOrders.OrderNo, dbo.FCustomerOrders.DT, dbo.FCustomerOrders.TradeTerms, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.PartialShipment, dbo.FCustomerOrders.PaymentTerms, dbo.FCustomerOrders.TransShipment, dbo.FCustomerOrders.Packaging, 
         dbo.FCustomerOrders.DeliveryDT, dbo.FCustomerOrders.BatchNo, dbo.FCustomerOrders.CompanyRefID, dbo.FCustomerOrders.StampDT, 
		 dbo.FCustomerOrders.Quality, dbo.FCustomerOrders.InternalRefNo, dbo.FCustomerOrders.OrderRcvdVia, dbo.FCustomerOrders.OrderType, TFOrderItems.SetItem, TFOrderItems.OrderItemID, TFOrderItems.CustomerItemCode, 
         TFOrderItems.CompItemCode, TFOrderItems.DeliveryDT AS DeliveryDTItem,TFOrderItems.Stamps AS FOI_Stamps,TFOrderItems.IW_BatchNo,
		 TFOrderItems.Finishing,TFOrderItems.ItemOrderNo,TFOrderItems.DeliveryDate,TFOrderItems.ItemRemarks
FROM  dbo.FCustomerOrders INNER JOIN
             (SELECT dbo.FOrderItems.OrderNo, dbo.FOrderItems.ItemCode, dbo.FOrderItems.CompItemCode, 0 AS SetItem, 
			 dbo.FOrderItems.CompItemCode AS OrderItemID, SUM(dbo.FOrderItems.Qty) AS Qty, dbo.FOrderItems.ItemCode AS CustomerItemCode, 
			 dbo.FOrderItems.DeliveryDT,dbo.FOrderItems.Stamps,IW_BatchNo AS IW_BatchNo,
			 Quality as Finishing,IW_OrderNo As ItemOrderNo,DeliveryDT AS DeliveryDate,Remarks as ItemRemarks
            FROM  dbo.FOrderItems INNER JOIN
                     dbo.Items ON dbo.Items.ItemID = dbo.FOrderItems.CompItemCode
            WHERE (dbo.Items.ItemType <> 2)
            GROUP BY dbo.FOrderItems.OrderNo, dbo.FOrderItems.ItemCode, dbo.FOrderItems.CompItemCode, dbo.FOrderItems.DeliveryDT,dbo.FOrderItems.Stamps,IW_BatchNo,
			Quality,IW_OrderNo,DeliveryDT,Remarks
            UNION ALL
            SELECT FOrderItems_1.OrderNo, dbo.ItemsSets.Set_ItemID, dbo.ItemsSets.Set_ItemID AS Expr1, 1 AS SetItem, FOrderItems_1.CompItemCode AS OrderItemID, 
			SUM(FOrderItems_1.Qty * dbo.ItemsSets.Qty) AS Qty, FOrderItems_1.ItemCode AS CustomerItemCode, FOrderItems_1.DeliveryDT,FOrderItems_1.Stamps
			,IW_BatchNo AS IW_BatchNo,Quality AS Finishing,IW_OrderNo As ItemOrderNo,DeliveryDT AS DeliveryDate,Remarks as ItemRemarks
            FROM  dbo.FOrderItems AS FOrderItems_1 INNER JOIN
                     dbo.Items AS Items_1 ON FOrderItems_1.CompItemCode = Items_1.ItemID INNER JOIN
                     dbo.ItemsSets ON dbo.ItemsSets.ItemID = Items_1.ItemID
            WHERE (Items_1.ItemType = 2)
            GROUP BY FOrderItems_1.OrderNo, dbo.ItemsSets.Set_ItemID, FOrderItems_1.CompItemCode, FOrderItems_1.ItemCode, FOrderItems_1.DeliveryDT,
			FOrderItems_1.Stamps,FOrderItems_1.IW_BatchNo,Quality,IW_OrderNo,DeliveryDT,Remarks
			) AS TFOrderItems ON dbo.FCustomerOrders.OrderNo = TFOrderItems.OrderNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRunningLots_POI_Old]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRunningLots_POI_Old];
END
GO

CREATE VIEW [dbo].[VRunningLots_POI_Old]
AS
SELECT     TUnion.StartDT, TUnion.LotNo, TUnion.ItemCode, TUnion.Qty, TUnion.Description, TUnion.LastProcessDays, TUnion.OrderNo, TUnion.Remarks, TUnion.DT, 
                      TUnion.ProcessID, TUnion.ReWorkLot, VItems_1.GroupID, FCustomerOrders_1.CustCode, FCustomerOrders_1.Country, TRunning.TLRunning, 
                      TRunningOrder.TLRunning_Order, TSF.R1, TSF.R2, TPack.PackStock, VItems_1.ItemName, VUnshippedOrderItemList_1.OrderNo AS UnShipped_OrderNo, 
                      TUnion.EntryType, TUnion.POIRemarks, TUnion.OpenPOQty, TUnion.VendID, TUnion.R1 AS RunningLots_R1, TLOrder.TLUnShippedQty, 
                      TLot_Wastage.Lot_Wastage_Qty, VUnshippedOrderItemList_1.Price, VUnshippedOrderItemList_1.PriceForCost, dbo.FOrderItems.Price AS FOI_Price, 
                      dbo.ForeignCustomers.Curr
FROM         dbo.ForeignCustomers INNER JOIN
                      dbo.FCustomerOrders AS FCustomerOrders_1 ON dbo.ForeignCustomers.CustCode = FCustomerOrders_1.CustCode AND 
                      dbo.ForeignCustomers.Country = FCustomerOrders_1.Country RIGHT OUTER JOIN
                          (SELECT     ISNULL(T1.StartDT, T2.StartDT) AS StartDT, dbo.RunningLots_Temp.LotNo, dbo.RunningLots_Temp.ItemCode, dbo.RunningLots_Temp.Qty, 
                                                   dbo.RunningLots_Temp.Description, DATEDIFF(day, dbo.RunningLots_Temp.DT, GETDATE()) AS LastProcessDays, dbo.RunningLots_Temp.OrderNo, 
                                                   '' AS Remarks, dbo.RunningLots_Temp.DT, dbo.RunningLots_Temp.ProcessID, dbo.RunningLots_Temp.ReWorkLot, 0 AS EntryType, 
                                                   dbo.POI_Remarks.Remarks AS POIRemarks, 0 AS OpenPOQty, dbo.RunningLots_Temp.VendID, 
                                                   CASE WHEN SNo <= PolishingItemSNo THEN 1 ELSE 0 END AS R1
                             FROM         dbo.RunningLots_Temp LEFT OUTER JOIN
                                                   dbo.POI_Remarks ON dbo.RunningLots_Temp.LotNo = dbo.POI_Remarks.LotNo LEFT OUTER JOIN
                                                       (SELECT     dbo.vendRcvdDetail.LotNo, MIN(dbo.VendReceived.DT) AS StartDT
                                                          FROM         dbo.VendReceived INNER JOIN
                                                                                dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID
                                                          GROUP BY dbo.vendRcvdDetail.LotNo) AS T1 ON dbo.RunningLots_Temp.LotNo = T1.LotNo INNER JOIN
                                                   dbo.VItems ON dbo.RunningLots_Temp.ItemCode = dbo.VItems.ItemID LEFT OUTER JOIN
                                                   dbo.FCustomerOrders ON dbo.RunningLots_Temp.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                                                       (SELECT     dbo.VendIssdDetail.LotNo, MIN(dbo.VendIssued.DT) AS StartDT
                                                          FROM         dbo.VendIssued INNER JOIN
                                                                                dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID
                                                          GROUP BY dbo.VendIssdDetail.LotNo) AS T2 ON dbo.RunningLots_Temp.LotNo = T2.LotNo
                             UNION ALL
                             SELECT     dbo.VUnshippedOrderItemList.DT, '' AS LotNo, dbo.VUnshippedOrderItemList.CompItemCode, 
                                                   dbo.VUnshippedOrderItemList.OrderQty - dbo.VUnshippedOrderItemList.ShippedQty - ISNULL(T1_1.RLQty, 0) AS Qty, '' AS Description, 
                                                   0 AS LastProcessDays, dbo.VUnshippedOrderItemList.OrderNo, dbo.VUnshippedOrderItemList.Remarks, dbo.VUnshippedOrderItemList.DT AS Expr1, 
                                                   0 AS ProcessID, 0 AS ReWorkLot, 1 AS EntryType, POI_Remarks_1.Remarks AS POIRemarks, TOpenPOs.OpenPOQty, 0 AS VendID, 0 AS R1
                             FROM         dbo.VUnshippedOrderItemList LEFT OUTER JOIN
                                                   dbo.POI_Remarks AS POI_Remarks_1 ON dbo.VUnshippedOrderItemList.OrderNo = POI_Remarks_1.OrderNo AND 
                                                   dbo.VUnshippedOrderItemList.CompItemCode = POI_Remarks_1.ItemCode LEFT OUTER JOIN
                                                       (SELECT     OrderNo, ItemCode, SUM(Qty) AS OpenPOQty
                                                          FROM         dbo.VItemOpenPOs
                                                          GROUP BY OrderNo, ItemCode) AS TOpenPOs ON dbo.VUnshippedOrderItemList.OrderNo = TOpenPOs.OrderNo AND 
                                                   dbo.VUnshippedOrderItemList.CompItemCode = TOpenPOs.ItemCode LEFT OUTER JOIN
                                                       (SELECT     OrderNo, ItemCode, SUM(Qty) AS RLQty
                                                          FROM         dbo.RunningLots_Temp AS VRunningLots_Simple_3
                                                          GROUP BY OrderNo, ItemCode) AS T1_1 ON dbo.VUnshippedOrderItemList.OrderNo = T1_1.OrderNo AND 
                                                   dbo.VUnshippedOrderItemList.CompItemCode = T1_1.ItemCode) AS TUnion INNER JOIN
                      dbo.VItems AS VItems_1 ON TUnion.ItemCode = VItems_1.ItemID INNER JOIN
                      dbo.FOrderItems ON TUnion.ItemCode = dbo.FOrderItems.CompItemCode AND TUnion.OrderNo = dbo.FOrderItems.OrderNo ON 
                      FCustomerOrders_1.OrderNo = TUnion.OrderNo LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(Qty) AS TLRunning
                             FROM         dbo.RunningLots_Temp AS VRunningLots_Simple_2
                             GROUP BY OrderNo, ItemCode) AS TRunning ON TUnion.OrderNo = TRunning.OrderNo AND TUnion.ItemCode = TRunning.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(Qty) AS TLRunning_Order
                             FROM         dbo.RunningLots_Temp AS VRunningLots_Simple_1
                             GROUP BY OrderNo, ItemCode) AS TRunningOrder ON TUnion.OrderNo = TRunningOrder.OrderNo AND 
                      TUnion.ItemCode = TRunningOrder.ItemCode LEFT OUTER JOIN
                          (SELECT     ItemID, SUM(CASE WHEN R1 = 1 THEN Qty - IssdQty ELSE 0 END) AS R1, SUM(CASE WHEN R1 = 0 THEN Qty - IssdQty ELSE 0 END) AS R2
                             FROM         dbo.VStockOrderOpening
                             WHERE     (Qty > IssdQty)
                             GROUP BY ItemID) AS TSF ON TUnion.ItemCode = TSF.ItemID LEFT OUTER JOIN
                          (SELECT     ItemID, SUM(NetQty) AS PackStock
                             FROM         dbo.VItemsWithShelfWiseStock
                             GROUP BY ItemID) AS TPack ON TUnion.ItemCode = TPack.ItemID LEFT OUTER JOIN
                      dbo.VUnshippedOrderItemList AS VUnshippedOrderItemList_1 ON TUnion.ItemCode = VUnshippedOrderItemList_1.CompItemCode AND 
                      TUnion.OrderNo = VUnshippedOrderItemList_1.OrderNo LEFT OUTER JOIN
                          (SELECT     CompItemCode, SUM(OrderQty - ShippedQty) AS TLUnShippedQty
                             FROM         dbo.VUnshippedOrderItemList AS VUnshippedOrderItemList_2
                             GROUP BY CompItemCode) AS TLOrder ON VItems_1.ItemID = TLOrder.CompItemCode LEFT OUTER JOIN
                          (SELECT     LotNo, SUM(Wastage) AS Lot_Wastage_Qty
                             FROM         dbo.vendRcvdDetail AS vendRcvdDetail_1
                             GROUP BY LotNo) AS TLot_Wastage ON TUnion.LotNo = TLot_Wastage.LotNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VSFStockTransaction]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VSFStockTransaction];
END
GO


-- ============================================================================
-- Fix VSFStockTransaction View: Ensure LotNo is populated for manual SF open receiving
-- Fallback to StockOrderOpening.LotNo_Manual when VendRcvdDetail.LotNo is null
-- ============================================================================
CREATE   VIEW [dbo].[VSFStockTransaction]
AS
SELECT        dbo.StockOrderOpening.EntryID, dbo.VendIssued.UserName, dbo.VendIssued.DT, dbo.VendIssdDetail.ItemCode, dbo.Items.ItemName, dbo.StockOrderOpening.ProcessID, dbo.Processes.Description, dbo.StockOrderOpening_Issuance.QtyIssd, 1 AS Issuance, 
                         Location, StoreName, Store_RefID, ISNULL(NULLIF(VendIssdDetail.LotNo, ''), StockOrderOpening.LotNo_Manual) AS LotNo
                         ,dbo.VendIssdDetail.OrderNo,StockOrderOpening.Remarks
FROM            dbo.StockOrderOpening_Issuance INNER JOIN
                         dbo.StockOrderOpening ON dbo.StockOrderOpening_Issuance.SOO_RefID = dbo.StockOrderOpening.EntryID INNER JOIN
                         dbo.VendIssdDetail ON dbo.VendIssdDetail.EntryID = dbo.StockOrderOpening_Issuance.VID_RefID INNER JOIN
                         dbo.VendIssued ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
                         dbo.Items ON dbo.VendIssdDetail.ItemCode = dbo.Items.ItemID INNER JOIN
                         dbo.Processes ON dbo.StockOrderOpening.ProcessID = dbo.Processes.ProcessID INNER JOIN
                         VStoreShelfs ON VStoreShelfs.EntryID = StockOrderOpening.Shelf_RefID
UNION
SELECT        StockOrderOpening.EntryID, UserName, CAST(CONVERT(VARCHAR(50), DTEntry, 6) AS DATETIME), Items.ItemID, ItemName, StockOrderOpening.ProcessID, Processes.Description, StockOrderOpening.Qty, 0 AS Issuance, Location, StoreName, Store_RefID, 
                         ISNULL(NULLIF(VendRcvdDetail.LotNo, ''), StockOrderOpening.LotNo_Manual) AS LotNo
                         ,VendRcvdDetail.OrderNo,StockOrderOpening.Remarks
FROM            StockOrderOpening INNER JOIN
                         Items ON StockOrderOpening.ItemID = Items.ItemID INNER JOIN
                         Processes ON StockOrderOpening.ProcessID = Processes.ProcessID INNER JOIN
                         VStoreShelfs ON VStoreShelfs.EntryID = StockOrderOpening.Shelf_RefID LEFT OUTER JOIN
                         LotTransferDetails ON StockOrderOpening.EntryID = LotTransferDetails.SOO_RefID LEFT OUTER JOIN
                         VendRcvdDetail ON LotTransferDetails.VRD_From_RefID = VendRcvdDetail.EntryID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VStockTransactions]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VStockTransactions];
END
GO


CREATE VIEW [dbo].[VStockTransactions]
AS
SELECT     StockTransactions.EntryNo, StockTransactions.DT, StockTransactions.OrderNo, StockTransactions.RcvdBy, StockTransactions.RcvdFrom, StockTransactions.ItemID, StockTransactions.Qty, StockTransactions.Location, StockTransactions.ChildEntryID, 
                  StockTransactions.EntryType, StockTransactions.Shelf_RefID, dbo.Items.ItemName, dbo.VStoreShelfs.Store_RefID, StockTransactions.SampleEntry, StockTransactions.LotNo_Manual, dbo.Items.ItemPic, StockTransactions.Remarks, dbo.VStoreShelfs.StoreName, 
                  dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, dbo.Items.ItemNameUrdu, dbo.Items.Unit,StockTransactions.TransferRemarks
FROM        dbo.Items INNER JOIN
                      (SELECT     dbo.RcvItemsSimple.RcvNo AS EntryNo, CAST(CONVERT(VARCHAR(50), dbo.RcvItemsSimple.DT, 6) AS DATETIME) AS DT, dbo.RcvItemsSimple.OrderNo, dbo.RcvItemsSimple.RcvdBy, dbo.RcvItemsSimple.RcvdFrom, dbo.RcvItemsSimpleDetail.ItemID, 
                                         dbo.RcvItemsSimpleDetail.Qty, dbo.RcvItemsSimpleDetail.Location, dbo.RcvItemsSimpleDetail.EntryID AS ChildEntryID, 1 AS EntryType, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID, 0 AS SampleEntry, dbo.RcvItemsSimpleDetail.LotNo_Manual, 
                                         dbo.RcvItemsSimpleDetail.Remarks,
										 dbo.RcvItemsSimpleDetail_Placement.Remarks TransferRemarks
                       FROM        dbo.RcvItemsSimple INNER JOIN
                                         dbo.RcvItemsSimpleDetail ON dbo.RcvItemsSimple.RcvNo = dbo.RcvItemsSimpleDetail.RcvNo INNER JOIN
                                         dbo.RcvItemsSimpleDetail_Placement ON dbo.RcvItemsSimpleDetail.EntryID = dbo.RcvItemsSimpleDetail_Placement.RISD_RefID
                       UNION
                       SELECT     dbo.IssItemsSimple.IssNo AS EntryNo, CAST(CONVERT(VARCHAR(50), dbo.IssItemsSimple.DT, 6) AS DATETIME) AS DT, dbo.IssItemsSimple.OrderNo, dbo.IssItemsSimple.Department, dbo.IssItemsSimple.AttnPerson, dbo.IssItemsSimpleDetail.ItemID, 
                                         dbo.IssItemsSimpleDetail_More.IssdQty AS Qty, dbo.IssItemsSimpleDetail.Location, dbo.IssItemsSimpleDetail.EntryID AS ChildEntryID, 0 AS EntryType, dbo.IssItemsSimpleDetail_More.Shelf_RefID, 
                                         dbo.IssItemsSimple.SampleIssuance AS SampleEntry, dbo.IssItemsSimpleDetail.LotNo_Manual, 
										 dbo.IssItemsSimpleDetail.Remarks,
										 dbo.IssItemsSimpleDetail_More.Remarks  TransferRemarks
                       FROM        dbo.IssItemsSimple INNER JOIN
                                         dbo.IssItemsSimpleDetail ON dbo.IssItemsSimple.IssNo = dbo.IssItemsSimpleDetail.IssNo INNER JOIN
                                         dbo.IssItemsSimpleDetail_More ON dbo.IssItemsSimpleDetail.EntryID = dbo.IssItemsSimpleDetail_More.IISD_RefID
				  ) AS StockTransactions ON dbo.Items.ItemID = StockTransactions.ItemID LEFT OUTER JOIN
                  dbo.VStoreShelfs ON StockTransactions.Shelf_RefID = dbo.VStoreShelfs.EntryID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VTransferredToReadyFinishLotsDetail]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VTransferredToReadyFinishLotsDetail];
END
GO


CREATE VIEW [dbo].[VTransferredToReadyFinishLotsDetail]
AS
SELECT dbo.vendRcvdDetail.LotNo, dbo.TransferredToReadyFinishLotsDetail.Qty, dbo.TransferredToReadyFinishLotsDetail.EntryID, dbo.TransferredToReadyFinishLotsDetail.RefID, dbo.TransferredToReadyFinishLotsDetail.VRD_RefID, dbo.TransferredToReadyFinishLots.DT, dbo.TransferredToReadyFinishLots.UserName, dbo.TransferredToReadyFinishLots.MachineName, 
         dbo.vendRcvdDetail.ItemCode, dbo.TransferredToReadyFinishLots.ItemID, dbo.TransferredToReadyFinishLots.Remarks
         
		 ,VRcvItemsSimpleDetail_Placement.Shelf_RefID,VRcvItemsSimpleDetail_Placement.StoreName,VRcvItemsSimpleDetail_Placement.ShelfNo,VRcvItemsSimpleDetail_Placement.RackNo
FROM  dbo.TransferredToReadyFinishLots INNER JOIN
         dbo.TransferredToReadyFinishLotsDetail ON dbo.TransferredToReadyFinishLots.EntryID = dbo.TransferredToReadyFinishLotsDetail.RefID INNER JOIN
         dbo.vendRcvdDetail ON dbo.TransferredToReadyFinishLotsDetail.VRD_RefID = dbo.vendRcvdDetail.EntryID
         LEFT JOIN VRcvItemsSimpleDetail_Placement ON dbo.TransferredToReadyFinishLots.EntryID=VRcvItemsSimpleDetail_Placement.TTRFLD_RefID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVenderAssItems]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVenderAssItems];
END
GO


CREATE VIEW [dbo].[VVenderAssItems]
AS
SELECT     dbo.VenderAssItems.VendID, dbo.VenderAssItems.RMID, dbo.VenderAssItems.Rate, dbo.RM.RMID1, dbo.RM.RMName, dbo.RM.Unit, 
                      dbo.RMGroups.Description AS GroupName, dbo.VenderAssItems.EntryID
                      ,dbo.VenderAssItems.Remarks
                      ,Accounts.AccTitle
FROM         dbo.VenderAssItems INNER JOIN
                      dbo.RM ON dbo.VenderAssItems.RMID = dbo.RM.RMID INNER JOIN
                      dbo.RMGroups ON dbo.RM.GroupID = dbo.RMGroups.ID
                      INNER JOIN Accounts ON dbo.VenderAssItems.VendID=Accounts.AccNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVenders]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVenders];
END
GO

CREATE   VIEW [dbo].[VVenders]
AS
SELECT        dbo.Venders.VendID, dbo.Venders.AccNo, dbo.Venders.Phone1, dbo.Venders.Phone2, dbo.Venders.Phone3, dbo.Venders.Fax1, dbo.Venders.Fax2, dbo.Venders.Address, dbo.Venders.ContactPerson, dbo.Venders.CPhone, 
                         dbo.Venders.CEmail, dbo.Venders.Mobile, dbo.Venders.ImportVender, dbo.Venders.MakerNo, dbo.Venders.BankAccNo, dbo.Venders.ProcessID, dbo.Venders.VenderNameUrdu, dbo.Venders.VenderDescription, 
                         dbo.Venders.VenderPic, dbo.Venders.VenderSig, dbo.Accounts.Balance, dbo.Accounts.openbal, dbo.Accounts.opendate, dbo.Accounts.AccTitle, dbo.Accounts.Active, dbo.Venders.MakerNo AS MKNo, dbo.Accounts.SubAccOf
FROM            dbo.Venders INNER JOIN
                         dbo.Accounts ON dbo.Venders.AccNo = dbo.Accounts.AccNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendersAndMakers]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendersAndMakers];
END
GO

CREATE   VIEW [dbo].[VVendersAndMakers]
AS
SELECT        T1.VendID, T1.AccNo, T1.AccTitle, T1.Phone1, T1.Phone2, T1.Fax1, T1.Address, T1.Active, T1.Vender, T1.MakerNo, T1.SubAccOf, dbo.Accounts.AccTitle AS AccTitle_Parent
FROM            (SELECT        VendID, AccNo, AccTitle, Phone1, Phone2, Fax1, Address, Active, 1 AS Vender, MakerNo, SubAccOf
                          FROM            dbo.VVenders
                          UNION
                          SELECT        VendID, AccNo, AccTitle, Phone1, Phone2, Fax1, Address, Active, 0 AS Vender, VendID1, SubAccOf
                          FROM            dbo.VMakers) AS T1 INNER JOIN
                         dbo.Accounts ON T1.SubAccOf = dbo.Accounts.AccNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendIssdDetail_ForRunningLots]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendIssdDetail_ForRunningLots];
END
GO


CREATE   VIEW [dbo].[VVendIssdDetail_ForRunningLots]
AS
SELECT        dbo.VendIssued.VendID, dbo.VendIssued.DT, dbo.VendIssued.UserID, dbo.VendIssued.ProcessID, dbo.VendIssdDetail.EntryID, dbo.VendIssdDetail.RefID, 
                         dbo.VendIssdDetail.RecieptID, dbo.VendIssdDetail.ItemCode, dbo.VendIssdDetail.Rate, dbo.VendIssdDetail.IssQty, dbo.VendIssdDetail.RcvdQty, 
                         dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail.ReqAuth, dbo.VendIssdDetail.OrderNo, dbo.VendIssdDetail.RcvProcessID, dbo.VendIssdDetail.Rcvd_RefID, 
                         dbo.VendIssdDetail.ReWorkLot, dbo.VendIssdDetail.Repair_RefID, dbo.VendIssued.RecieptID AS MainRecieptID, dbo.Processes.Description, dbo.Processes.Code, 
                         dbo.ItemProcesses.SNO, dbo.VendIssued.Closed, dbo.VendIssdDetail.ReturnDT, dbo.VendIssued.MasterPONo, dbo.FCustomerOrders.InternalRefNo
                         ,dbo.VendIssued.IssEmpID
FROM            dbo.VendIssued INNER JOIN
                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
                         dbo.Processes ON dbo.VendIssued.ProcessID = dbo.Processes.ProcessID INNER JOIN
                         dbo.FCustomerOrders ON dbo.VendIssdDetail.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                         dbo.ItemProcesses ON dbo.VendIssued.ProcessID = dbo.ItemProcesses.ProcessID AND dbo.VendIssdDetail.ItemCode = dbo.ItemProcesses.ItemID


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendIssdDetailWithValueWORcving]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendIssdDetailWithValueWORcving];
END
GO


CREATE   VIEW [dbo].[VVendIssdDetailWithValueWORcving]
AS
SELECT        dbo.VendIssued.EntryID AS MainEntryID, dbo.VendIssued.RecieptID AS MainRecieptID, dbo.VendIssued.DT, dbo.VendIssdDetail.EntryID, dbo.VendIssued.VendID, 
                         dbo.VendIssdDetail.RefID, dbo.VendIssdDetail.RecieptID, dbo.VendIssdDetail.ItemCode, dbo.VendIssdDetail.Rate, dbo.VendIssdDetail.IssQty, 
                         dbo.VendIssdDetail.RcvdQty, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail.ReqAuth, dbo.VendIssdDetail.OrderNo, dbo.VendIssdDetail.RcvProcessID, 
                         dbo.VendIssdDetail.ReturnDT, dbo.VendIssdDetail.Priority, dbo.VendIssdDetail.SpecialInstructions, dbo.VendIssued.VchrNo, dbo.Items.ForgingWeight, 
                         VendIssdDetail.IssQty * VendIssdDetail.Rate AS IssValue, 
                         (VendIssdDetail.IssQty - VendIssdDetail.RcvdQty) * VendIssdDetail.Rate AS BalanceValue
                         , dbo.VendIssdDetail.ReWorkLot, 
                         dbo.VendIssdDetail.Repair_RefID, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.InternalRefNo, dbo.VID_Bookmarks.EntryID AS BookMarkEntryID, 
                         T1.ItemCode AS CustItemCode, 
			T1.Quality, dbo.VIssuanceEntryWithWastage.WastageQty,T1.OrderQty,VendIssdDetail.Batch_No
FROM            dbo.VendIssdDetail INNER JOIN
                         dbo.VendIssued ON dbo.VendIssdDetail.RefID = dbo.VendIssued.EntryID INNER JOIN
                         dbo.Items ON dbo.VendIssued.ItemID = dbo.Items.ItemID INNER JOIN
                         dbo.FCustomerOrders ON dbo.VendIssdDetail.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                         dbo.VID_Bookmarks ON dbo.VendIssdDetail.EntryID = dbo.VID_Bookmarks.VID_RefID LEFT OUTER JOIN
                             (SELECT        OrderNo, ItemCode,CompItemCode, MAX(Quality) AS Quality,SUM(Qty) AS OrderQty
                                FROM            dbo.FOrderItems
                                GROUP BY OrderNo, ItemCode,CompItemCode) AS T1 ON dbo.VendIssdDetail.OrderNo = T1.OrderNo AND 
                         dbo.VendIssdDetail.ItemCode = T1.CompItemCode LEFT OUTER JOIN
                         dbo.VIssuanceEntryWithWastage ON dbo.VendIssdDetail.EntryID = dbo.VIssuanceEntryWithWastage.EntryID


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendIssued]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendIssued];
END
GO





CREATE   VIEW [dbo].[VVendIssued]
AS
SELECT     dbo.VendIssued.VendID, dbo.VendIssued.DT, dbo.VendIssued.RecieptID, dbo.VendIssued.ItemID, dbo.Makers.VendID1, dbo.Makers.VenderName, CAST(CONVERT(Varchar(10), dbo.VendIssued.DT, 1) AS DATETIME) AS OnlyDT, TabValue.TotalValue, 
                  dbo.VendIssued.Authorized, dbo.MakerPOPromises_F(dbo.VendIssued.EntryID) AS MakerPOReturnDTs, dbo.Processes.Description, dbo.Items.ItemName, TabValue.TotalIssQty, TabValue.LotNo, dbo.VendIssued.MasterPONo, TabValue.RcvdQty, dbo.Items.CatID, 
                  TabValue.ReWorkLot, dbo.VendIssued.ProcessID, dbo.Items.GroupID, TabValue.OrderNo, dbo.VendIssued.Closed, dbo.FCustomerOrders.CustCode, TabValue.LastRcvDT, TabValue.EntryID AS DetailEntryID, dbo.FCustomerOrders.InternalRefNo, 
                  TabValue.BookMarkEntryID, dbo.VendIssued.EntryID, dbo.Items.TipSize, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Makers.Phone1, dbo.Makers.Phone2, dbo.Makers.CPhone, dbo.Makers.ContactPerson, dbo.Makers.Address, dbo.Makers.Mobile, 
                  dbo.Makers.CompanyName, dbo.Makers.Maker_Second_Name, dbo.Makers.MakerNameUrdu, dbo.Makers.Fax1,
				  dbo.VendIssued.IssEmpID,dbo.VEmp.Name AS Emp_Name,VendIssued.DTEntry
                  ,TabValue.Batch_No
                  ,VHubSupervisors.Hub_Name,VHubSupervisors.Supervisors
FROM        dbo.VendIssued INNER JOIN
                  dbo.Items ON dbo.VendIssued.ItemID = dbo.Items.ItemID INNER JOIN
                  dbo.Processes ON dbo.VendIssued.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
                  dbo.Makers ON dbo.VendIssued.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                  dbo.FCustomerOrders RIGHT OUTER JOIN
                      (SELECT     dbo.VVendIssdDetailWithValueWORcving.OrderNo, dbo.VVendIssdDetailWithValueWORcving.EntryID, dbo.VVendIssdDetailWithValueWORcving.RefID, dbo.VVendIssdDetailWithValueWORcving.ReWorkLot, 
                                         dbo.VVendIssdDetailWithValueWORcving.BookMarkEntryID, SUM(dbo.VVendIssdDetailWithValueWORcving.IssValue) AS TotalValue, SUM(dbo.VVendIssdDetailWithValueWORcving.IssQty) AS TotalIssQty, 
                                         MAX(dbo.VVendIssdDetailWithValueWORcving.LotNo) AS LotNo, SUM(TRcv.RcvRcvdQty) AS RcvdQty, MAX(TRcv.RcvDT) AS LastRcvDT
                                         ,dbo.VVendIssdDetailWithValueWORcving.Batch_No
                       FROM        dbo.VVendIssdDetailWithValueWORcving LEFT OUTER JOIN
                                             (SELECT     dbo.vendRcvdDetail.Issue_RefID, MAX(dbo.VendReceived.DT) AS RcvDT, SUM(dbo.vendRcvdDetail.RcvdQty) AS RcvRcvdQty
                                              FROM        dbo.VendReceived INNER JOIN
                                                                dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID
                                              GROUP BY dbo.vendRcvdDetail.Issue_RefID) AS TRcv ON dbo.VVendIssdDetailWithValueWORcving.EntryID = TRcv.Issue_RefID
                       GROUP BY dbo.VVendIssdDetailWithValueWORcving.OrderNo, dbo.VVendIssdDetailWithValueWORcving.EntryID, dbo.VVendIssdDetailWithValueWORcving.RefID, dbo.VVendIssdDetailWithValueWORcving.ReWorkLot, 
                                         dbo.VVendIssdDetailWithValueWORcving.BookMarkEntryID,dbo.VVendIssdDetailWithValueWORcving.Batch_No) AS TabValue ON dbo.FCustomerOrders.OrderNo = TabValue.OrderNo ON dbo.VendIssued.EntryID = TabValue.RefID
				LEFT JOIN VEmp ON dbo.VendIssued.IssEmpID=VEmp.EmpID
                LEFT OUTER JOIN dbo.ItemProcessGroups ON VendIssued.ItemID = dbo.ItemProcessGroups.ItemID 
                LEFT OUTER JOIN dbo.ProcessGroupsProcesses ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroupsProcesses.Group_RefID 
                                                          AND dbo.VendIssued.ProcessID = dbo.ProcessGroupsProcesses.Process_RefID 
                LEFT OUTER JOIN dbo.VHubSupervisors ON dbo.ProcessGroupsProcesses.Group_RefID = dbo.VHubSupervisors.GroupID 
                                                  AND dbo.ProcessGroupsProcesses.Hub_Name = dbo.VHubSupervisors.Hub_Name
                                                  

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendIssued_rpt]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendIssued_rpt];
END
GO

CREATE   VIEW dbo.VVendIssued_rpt
AS
SELECT     dbo.VendIssued.EntryID, dbo.VendIssued.VendID, dbo.VendIssued.DT, dbo.VendIssued.RecieptID, dbo.VendIssued.UserID, dbo.VendIssued.ProcessID, dbo.VendIssued.ItemID, dbo.VendIssued.UserName, dbo.VendIssued.MachineName, 
                  dbo.VendIssued.SpecialInstructions, dbo.VendIssued.VchrNo, dbo.VendIssued.Authorized, dbo.VendIssued.ExcessQtyPercentage, dbo.VendIssued.MaximumRcvingsAgainstPO, dbo.VendIssued.ProcessOrderNo, dbo.VendIssued.MasterPONo, 
                  dbo.VendIssued.SampleProvided, dbo.VendIssued.DrawingProvided, dbo.VendIssued.ForgingProvided, dbo.VendIssued.SteelProvided, dbo.VendIssued.Closed, dbo.VendIssued.IssEmpID, dbo.VendIssued.SteelType_RefID, dbo.VendIssued.AuthUserName, 
                  dbo.VendIssued.AuthMachineName, dbo.VendIssued.AuthEntryDT, ISNULL(T1.Posted, 0) AS Posted
FROM        dbo.VendIssued LEFT OUTER JOIN
                      (SELECT     VI_RefID, SUM(CASE WHEN ISNULL(VI_RefID, 0) > 0 THEN 1 ELSE 0 END) AS Posted
                       FROM        dbo.RawMaterialIssuance
                       GROUP BY VI_RefID) AS T1 ON dbo.VendIssued.EntryID = T1.VI_RefID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendRcvdDetail]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendRcvdDetail];
END
GO



CREATE VIEW [dbo].[VVendRcvdDetail]
AS
SELECT     dbo.VendReceived.DT, dbo.VendReceived.VendID, dbo.VendRcvdDetail.RcvdQty - dbo.VendRcvdDetail.Wastage - ISNULL(dbo.VendRcvdDetail.RcvRepairQty, 0) 
                      - dbo.VendRcvdDetail.IssQty AS IssuableQty, dbo.VendReceived.UserID, dbo.VItems1.CompleteItemName AS ItemDescription, dbo.VendReceived.Issuance_RefID, 
                      dbo.VendReceived.ProcessID AS MainProcessID, dbo.VendAssItems.Unit, dbo.VItems1.ForgingWeight, 
                      CASE WHEN VendAssItems.UNIT = 'Kgs' THEN (RcvdWeight - ((RcvdWeight / dbo.VendRcvdDetail.RcvdQty) * (Wastage + LostQty))) 
                      * VendIssdDetail.Rate ELSE VendIssdDetail.Rate * (dbo.VendRcvdDetail.RcvdQty - Wastage - LostQty) END AS RcvdValue, 
                      (CASE WHEN dbo.VendRcvdDetail.RcvdQty = 0 THEN 0 ELSE RcvdWeight / dbo.VendRcvdDetail.RcvdQty END) * dbo.VendRcvdDetail.LostQty AS LostQtyWeight, 
                      (CASE WHEN dbo.VendRcvdDetail.RcvdQty = 0 THEN 0 ELSE RcvdWeight / dbo.VendRcvdDetail.RcvdQty END) * dbo.VendRcvdDetail.Wastage AS WasteQtyWeight, 
                      dbo.VendIssdDetail.Rate AS IssRate, CAST(CONVERT(VARCHAR(50), dbo.VendReceived.DT, 6) AS DATETIME) AS OnlyDT, dbo.Processes.BillingProcessID, 
                      dbo.Makers.VenderName, dbo.Processes.Code, dbo.Processes.Description, dbo.VItems1.ItemName, dbo.VItems1.ItemSize, dbo.VItems1.SizeUnit, 
                      dbo.VItems1.TipSize, dbo.VendIssued.RecieptID AS MainRecieptID, dbo.VendReceived.EmpID, 
                      dbo.Employees.name + ' ' + dbo.Employees.Rel + ' ' + dbo.Employees.fname AS EmployeeName, dbo.ItemProcesses.SNO AS ItemSNo, 
                      ItemBillingProcess.SNO AS BillingProcessSNo, dbo.VendReceived.EntryID AS ReceivedEntryID, dbo.VendRcvdDetail.EntryID, dbo.VendRcvdDetail.RefID, 
                      dbo.VendRcvdDetail.ItemCode, dbo.VendRcvdDetail.RecieptID, dbo.VendRcvdDetail.RcvdQty, dbo.VendRcvdDetail.Wastage, dbo.VendRcvdDetail.Rate, 
                      dbo.VendRcvdDetail.LotNo, dbo.VendRcvdDetail.ReqAuth, dbo.VendRcvdDetail.RcvRepairQty, dbo.VendRcvdDetail.IssRepairQty, dbo.VendRcvdDetail.NextProcessID, 
                      dbo.VendRcvdDetail.LostQty, dbo.VendRcvdDetail.OrderNo, dbo.VendRcvdDetail.CountedBy, dbo.VendRcvdDetail.Issue_RefID, dbo.VendRcvdDetail.ProcessID, 
                      dbo.VendRcvdDetail.RcvdWeight, dbo.VendRcvdDetail.RepairAmt, dbo.VendRcvdDetail.Opening_RefID, dbo.VendRcvdDetail.ReWorkQty, 
                      dbo.VendRcvdDetail.ReWorkLot, dbo.VendRcvdDetail.Repair_RefID, dbo.VendRcvdDetail.Insp_EmpID, dbo.VendRcvdDetail.Audited, dbo.VendRcvdDetail.IssQty, 
                      dbo.VendAssItems.Rate AS VAIRate
                      ,VendIssdDetail.Rate * (dbo.VendRcvdDetail.RcvdQty) AS RcvdValue_For_Billing
FROM         dbo.VendReceived INNER JOIN
                      dbo.VendRcvdDetail ON dbo.VendReceived.EntryID = dbo.VendRcvdDetail.RefID INNER JOIN
                      dbo.VItems1 ON dbo.VendRcvdDetail.ItemCode = dbo.VItems1.ItemID 
					  LEFT JOIN FCustomerOrders ON VendRcvdDetail.OrderNo=FCustomerOrders.OrderNo
                      LEFT OUTER JOIN dbo.VendAssItems ON dbo.VendRcvdDetail.ItemCode = dbo.VendAssItems.ItemID AND dbo.VendReceived.VendID = dbo.VendAssItems.VendID AND dbo.VendRcvdDetail.ProcessID = dbo.VendAssItems.ProcessID  AND dbo.FCustomerOrders.CustCode=dbo.VendAssItems.CustCode
                      INNER JOIN dbo.VendIssdDetail ON dbo.VendIssdDetail.EntryID = dbo.VendRcvdDetail.Issue_RefID INNER JOIN
                      dbo.Processes ON dbo.VendRcvdDetail.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
                      dbo.ItemProcesses ON dbo.VendRcvdDetail.ItemCode = dbo.ItemProcesses.ItemID AND 
                      dbo.VendRcvdDetail.ProcessID = dbo.ItemProcesses.ProcessID LEFT OUTER JOIN
                      dbo.ItemProcesses AS ItemBillingProcess ON dbo.Processes.BillingProcessID = ItemBillingProcess.ProcessID AND 
                      dbo.VendRcvdDetail.ItemCode = ItemBillingProcess.ItemID INNER JOIN
                      dbo.Makers ON dbo.VendReceived.VendID = dbo.Makers.VendID INNER JOIN
                      dbo.VendIssued ON dbo.VendIssdDetail.RefID = dbo.VendIssued.EntryID LEFT OUTER JOIN
                      dbo.Employees ON dbo.VendReceived.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                          (SELECT     VRD_From_RefID, SUM(SplitQty) AS SplitQtyForRcvdQty
                             FROM         dbo.LotTransferDetails
                             GROUP BY VRD_From_RefID) AS TSplitQtyForRcvdQty ON dbo.VendRcvdDetail.EntryID = TSplitQtyForRcvdQty.VRD_From_RefID


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendRcvdDetail_QC_Report_PIP]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendRcvdDetail_QC_Report_PIP];
END
GO

CREATE   VIEW [dbo].[VVendRcvdDetail_QC_Report_PIP]
AS
SELECT dbo.VendRcvdDetail_QC_Report_PIP.VRD_RefID, dbo.VendRcvdDetail_QC_Report_PIP.PIP_RefID, dbo.VendRcvdDetail_QC_Report_PIP.DT, dbo.VendRcvdDetail_QC_Report_PIP.EmpID, dbo.VendRcvdDetail_QC_Report_PIP.Pieces, dbo.VendRcvdDetail_QC_Report_PIP.Remarks, dbo.VendRcvdDetail_QC_Report_PIP.Action, dbo.Process_Inspection_Points.Point_Description, 
         dbo.VEmp.name, dbo.VEmp.Designation, dbo.VEmp.DeptName
FROM  dbo.VendRcvdDetail_QC_Report_PIP INNER JOIN
         dbo.Process_Inspection_Points ON dbo.VendRcvdDetail_QC_Report_PIP.PIP_RefID = dbo.Process_Inspection_Points.EntryID LEFT OUTER JOIN
         dbo.VEmp ON dbo.VendRcvdDetail_QC_Report_PIP.EmpID = dbo.VEmp.empid

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendRcvdDetail_Simple]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendRcvdDetail_Simple];
END
GO

CREATE   VIEW [dbo].[VVendRcvdDetail_Simple]
AS
SELECT dbo.vendRcvdDetail.ItemCode, dbo.vendRcvdDetail.LotNo, dbo.vendRcvdDetail.OrderNo, dbo.Processes.Description, dbo.vendRcvdDetail.RcvdQty, dbo.vendRcvdDetail.IssQty, dbo.vendRcvdDetail.ProcessID, dbo.Processes.Code, dbo.ItemProcesses.SNO, dbo.vendRcvdDetail.Issue_RefID, dbo.vendRcvdDetail.EntryID, dbo.vendRcvdDetail.Wastage, dbo.vendRcvdDetail.ReWorkQty, 
         dbo.vendRcvdDetail.Opening_RefID, dbo.Items.GroupID, CAST(CONVERT(Varchar(10), dbo.VendReceived.DT, 1) AS DATETIME) AS DT, dbo.VendReceived.VendID, dbo.vendRcvdDetail.ReWorkLot, dbo.FCustomerOrders.InternalRefNo, dbo.FCustomerOrders.CustCode, dbo.VendReceived.OverTime, dbo.vendRcvdDetail.RefID, dbo.VendReceived.Issuance_RefID,
         dbo.vendRcvdDetail.NextProcessID
FROM  dbo.vendRcvdDetail INNER JOIN
         dbo.Processes ON dbo.vendRcvdDetail.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
         dbo.ItemProcesses ON dbo.vendRcvdDetail.ProcessID = dbo.ItemProcesses.ProcessID AND dbo.vendRcvdDetail.ItemCode = dbo.ItemProcesses.ItemID INNER JOIN
         dbo.Items ON dbo.vendRcvdDetail.ItemCode = dbo.Items.ItemID INNER JOIN
         dbo.VendReceived ON dbo.vendRcvdDetail.RefID = dbo.VendReceived.EntryID INNER JOIN
         dbo.FCustomerOrders ON dbo.vendRcvdDetail.OrderNo = dbo.FCustomerOrders.OrderNo

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVisitors]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVisitors];
END
GO

CREATE   VIEW [dbo].[VVisitors]
AS
SELECT
    v.EntryID,
    v.VisitorName,
    v.CompanyName,
    v.NICNumber,
    v.EmpID,
    v.DT,
    v.Purpose,
    v.Remarks,
    v.UserName,
    v.MachineName,
    v.DTEntry,
    d.DeptID AS deptid,
    d.Name AS DeptName,
    ISNULL(e.Name, v.EmpNameManual) AS Name,
    CAST(NULL AS varchar(100)) AS FName,
    CAST(NULL AS varchar(100)) AS Designation,
    CAST(CONVERT(date, v.DT) AS datetime) AS OnlyDT,
    v.TimeOut,
    v.ConcernedPerson,
    v.MobileNo,
    v.CardNo,
    v.VendID,
    CAST('' AS varchar(50)) AS VendCountry,
    ISNULL(vac.AccTitle, v.VendNameManual) AS VendName,
    v.EmpNameManual,
    v.VendNameManual,
    v.ConcernedPersonEmail,
    v.ConcernedPersonWhatsApp,
    v.NotifyByEmail,
    v.NotifyByWhatsApp,
    v.VisitCode,
    v.QrPayload,
    v.LastNotificationAt,
    v.NotificationStatus,
    v.OutVerifiedBy,
    v.OutVerifiedAt,
    v.OutVerificationInput
FROM dbo.Visitors v
LEFT JOIN dbo.Employees e ON e.empid COLLATE DATABASE_DEFAULT = v.EmpID COLLATE DATABASE_DEFAULT
LEFT JOIN dbo.Departments d ON d.deptid COLLATE DATABASE_DEFAULT = e.deptid COLLATE DATABASE_DEFAULT
LEFT JOIN dbo.Venders vv ON vv.VendID = v.VendID
LEFT JOIN dbo.Accounts vac ON vv.AccNo = vac.AccNo;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[vw_WorkflowApprovalHistory]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[vw_WorkflowApprovalHistory];
END
GO

-- ====================================================================================================
-- SCRIPT: Consolidated Workflow Approval History View
-- ====================================================================================================
CREATE   VIEW [dbo].[vw_WorkflowApprovalHistory]
AS
SELECT 
    'ProductionDeletionRequests' AS SourceTable,
    p.[Id] AS RecordId,
    p.[TaskId],
    CASE 
        WHEN p.[RequestType] = 'LotReceiving' THEN 'DeleteLotAuthorization'
        WHEN p.[RequestType] IN ('LotIssuance', 'MasterPOIssuance') THEN 'DeleteProductionIssuance'
        WHEN p.[RequestType] = 'SkipProcess' THEN 'SkipProcess'
        ELSE p.[RequestType]
    END AS WorkflowCode,
    ISNULL(wc.[WorkflowName], p.[RequestType]) AS WorkflowName,
    ISNULL(wc.[Module], 'Production') AS Module,
    p.[LotNo],
    p.[OrderNo],
    p.[ItemCode],
    p.[ItemName],
    p.[ProcessName],
    CASE 
        WHEN p.[RequestType] = 'SkipProcess' THEN NULL 
        ELSE p.[MakerName] 
    END AS MakerName,
    CASE 
        WHEN p.[RequestType] = 'SkipProcess' THEN p.[MakerName] 
        ELSE NULL 
    END AS TargetProcessName,
    p.[Qty],
    p.[Reason] AS RequestReason,
    p.[RequestedBy],
    p.[RequestedDT],
    p.[MachineName],
    p.[Status],
    p.[ReviewedBy] AS ActionBy,
    p.[ReviewedDT] AS ActionDT,
    p.[DirectorRemarks] AS ActionRemarks
FROM [dbo].[ProductionDeletionRequests] p WITH (NOLOCK)
LEFT JOIN [dbo].[WorkflowConfigurations] wc WITH (NOLOCK) ON wc.[WorkflowCode] = (
    CASE 
        WHEN p.[RequestType] = 'LotReceiving' THEN 'DeleteLotAuthorization'
        WHEN p.[RequestType] IN ('LotIssuance', 'MasterPOIssuance') THEN 'DeleteProductionIssuance'
        WHEN p.[RequestType] = 'SkipProcess' THEN 'SkipProcess'
        ELSE p.[RequestType]
    END
)

UNION ALL

SELECT 
    'TaskItems' AS SourceTable,
    t.[Id] AS RecordId,
    t.[Id] AS TaskId,
    CASE 
        WHEN t.[SourceEntityType] IN ('AttendanceWorkflow', 'AttendanceApproval') THEN 'AttendanceApproval'
        WHEN t.[SourceEntityType] IN ('MakerItemRateWorkflow', 'MakerRateChange') THEN 'MakerRateChange'
        ELSE t.[SourceEntityType]
    END AS WorkflowCode,
    ISNULL(wc.[WorkflowName], t.[SourceEntityType]) AS WorkflowName,
    ISNULL(wc.[Module], 'General') AS Module,
    NULL AS LotNo,
    NULL AS OrderNo,
    NULL AS ItemCode,
    t.[Title] AS ItemName,
    NULL AS ProcessName,
    NULL AS MakerName,
    NULL AS TargetProcessName,
    NULL AS Qty,
    t.[Description] AS RequestReason,
    t.[AssignedBy] AS RequestedBy,
    t.[CreatedAt] AS RequestedDT,
    NULL AS MachineName,
    CASE t.[Status]
        WHEN 2 THEN 'Approved'
        WHEN 3 THEN 'Rejected'
        WHEN 0 THEN 'Pending'
        WHEN 1 THEN 'Pending'
        ELSE 'Pending'
    END AS Status,
    t.[CompletedBy] AS ActionBy,
    t.[CompletedAt] AS ActionDT,
    (SELECT TOP 1 tc.[Content] FROM [dbo].[TaskComments] tc WITH (NOLOCK) WHERE tc.[TaskId] = t.[Id] ORDER BY tc.[Id] DESC) AS ActionRemarks
FROM [dbo].[TaskItems] t WITH (NOLOCK)
INNER JOIN [dbo].[WorkflowConfigurations] wc WITH (NOLOCK) ON wc.[WorkflowCode] = (
    CASE 
        WHEN t.[SourceEntityType] IN ('AttendanceWorkflow', 'AttendanceApproval') THEN 'AttendanceApproval'
        WHEN t.[SourceEntityType] IN ('MakerItemRateWorkflow', 'MakerRateChange') THEN 'MakerRateChange'
        ELSE t.[SourceEntityType]
    END
)
WHERE t.[SourceEntityType] NOT IN ('DeleteLotAuthorization', 'DeleteProductionIssuance', 'SkipProcess');

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[F_CustomInvoice_HSCodes]') AND type IN (N'FN', N'IF', N'TF'))
BEGIN
    DROP FUNCTION [dbo].[F_CustomInvoice_HSCodes];
END
GO

CREATE   FUNCTION [dbo].[F_CustomInvoice_HSCodes] (@CustomInvoice as Varchar(50))  
RETURNS Varchar(1000) AS

Begin
Declare @HSCodes As Varchar(1000)
Set @HSCodes=''
Declare @Temp As Varchar(255)
Declare Abc Cursor For Select Distinct HSCode From VrptCustomInvoiceDetail Where CustomInvoice=@CustomInvoice
Open Abc

Fetch Next From Abc Into @Temp
While @@Fetch_Status=0
	Begin
		Set @HSCodes=@HSCodes + ',' + @Temp
		Fetch Next From Abc Into @Temp
	End
	IF LEN(@HSCodes)>0 
		Set @HSCodes = Right(@HSCodes,Len(@HSCodes)-1)
Close Abc
Deallocate Abc

Return @HSCodes

End

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VGatePass_Return]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VGatePass_Return];
END
GO

CREATE   VIEW [dbo].[VGatePass_Return]
AS
SELECT        dbo.GatePass_Return.EntryID, dbo.GatePass_Return.GatePassNo, dbo.GatePass_Return.DT, dbo.GatePass_Return.PersonName, dbo.GatePass_Return.VehicleNo, 
              dbo.GatePass_Return.Description, dbo.GatePass_Return.UserName, dbo.GatePass_Return.MachineName, dbo.GatePass_Return.DTEntry, 
              CAST(CONVERT(VARCHAR(50), dbo.GatePass_Return.DT, 6) AS DATETIME) AS OnlyDT, dbo.GatePass_Return.SenderName, dbo.GatePass_Return.Returnable, 
              dbo.GatePass_Return.IssdTo, dbo.GatePass_Return.GatePassType, dbo.GatePass_Return.ReturnDT, T1.TotalQty, T1.TotalQtyRcvd, 
              dbo.GatePass_Return.DeptID_From, dbo.GatePass_Return.DeptID_To, dbo.Departments.name AS DeptName_From, 
              CASE WHEN DeptID_To = '0' THEN 'Outdoor Issuance' ELSE Departments_1.name END AS DeptName_To, dbo.GatePass_Return.VendID, 
              dbo.VVendersAndMakers.AccTitle AS Name, 
              dbo.GatePass_Return.ApprovedBy_Name, dbo.GatePass_Return.ApprovedBy_Designation,
              dbo.GatePass_Return.VM_AccNo
FROM          dbo.Departments RIGHT OUTER JOIN
              dbo.GatePass_Return LEFT OUTER JOIN
              dbo.VVendersAndMakers ON dbo.GatePass_Return.VM_AccNo = dbo.VVendersAndMakers.AccNo LEFT OUTER JOIN
              dbo.Departments AS Departments_1 ON dbo.GatePass_Return.DeptID_To = Departments_1.deptid ON 
              dbo.Departments.deptid = dbo.GatePass_Return.DeptID_From LEFT OUTER JOIN
              (SELECT MEGO_RefID, SUM(Qty) AS TotalQty, SUM(TotalQtyRcvd) AS TotalQtyRcvd
               FROM dbo.VGatePass_Return_Details
               GROUP BY MEGO_RefID) AS T1 ON dbo.GatePass_Return.EntryID = T1.MEGO_RefID;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRunningLots_Simple_Main]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRunningLots_Simple_Main];
END
GO





CREATE   VIEW [dbo].[VRunningLots_Simple_Main]
AS
SELECT        TOP (100) PERCENT T1.ItemCode, T1.LotNo, T1.OrderNo, T1.Description, T1.Qty, T1.ProcessID, T1.VendID, T1.ReWorkLot, dbo.VLotWithR2InDate.DT AS R2InDT, dbo.VItemPolishingProcessSNo.PolishingItemSNo, 
                         dbo.ItemProcesses.SNO, T1.DT, T1.LotType,T1.VRD_EntryID,T1.Lot_Available_On_ProcessID
                         ,T1.IssEmpID
FROM            (SELECT        ItemCode, LotNo, OrderNo, Description, IssQty - RcvdQty AS Qty, ProcessID, VendID, ReWorkLot, DT, 0 AS LotType,0 AS VRD_EntryID,RcvProcessID AS Lot_Available_On_ProcessID
                                ,dbo.VVendIssdDetail_ForRunningLots.IssEmpID
                          FROM            dbo.VVendIssdDetail_ForRunningLots
                          WHERE        (LotNo <> '0') AND (IssQty - RcvdQty > 0) AND (EntryID NOT IN
                                                        (SELECT        Issue_RefID
                                                          FROM            dbo.VendRcvdDetail))
                          UNION
                          SELECT        ItemCode, LotNo, OrderNo, Description, RcvdQty - IssQty - ISNULL(Wastage, 0) - ISNULL(ReWorkQty, 0) AS Qty, ProcessID, VendID, ReWorkLot, DT, 1 AS LotType,dbo.VVendRcvdDetail_Simple.EntryID AS VRD_EntryID,NextProcessID  AS Lot_Available_On_ProcessID
                                ,'' AS IssEmpID
                          FROM            dbo.VVendRcvdDetail_Simple
                          WHERE        (LotNo <> '0') AND (RcvdQty - IssQty - ISNULL(Wastage, 0) - ISNULL(ReWorkQty, 0) > 0) AND (EntryID NOT IN
                                                       (SELECT        Rcvd_RefID
                                                         FROM            dbo.VendIssdDetail)) AND (ISNULL(Opening_RefID, 0) = 0)) AS T1 LEFT OUTER JOIN
                         dbo.VLotWithR2InDate ON T1.LotNo = dbo.VLotWithR2InDate.LotNo LEFT OUTER JOIN
                         dbo.VItemPolishingProcessSNo ON T1.ItemCode = dbo.VItemPolishingProcessSNo.ItemCode LEFT OUTER JOIN
                         dbo.ItemProcesses ON T1.ItemCode = dbo.ItemProcesses.ItemID AND dbo.ItemProcesses.ProcessID = T1.ProcessID
WHERE        (T1.LotNo NOT IN
                             (SELECT        LotNo
                               FROM            dbo.Lots_Closed))
ORDER BY T1.ProcessID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRunningLots_Simple_WPP]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRunningLots_Simple_WPP];
END
GO

CREATE   VIEW [dbo].[VRunningLots_Simple_WPP]
AS
SELECT        dbo.VRunningLots_Simple_Main.ItemCode, dbo.VRunningLots_Simple_Main.LotNo, dbo.VRunningLots_Simple_Main.OrderNo, dbo.VRunningLots_Simple_Main.Description
						,dbo.VRunningLots_Simple_Main.Qty,dbo.VRunningLots_Simple_Main.ProcessID, dbo.VRunningLots_Simple_Main.VendID, dbo.VRunningLots_Simple_Main.ReWorkLot
						,dbo.VRunningLots_Simple_Main.R2InDT, dbo.VRunningLots_Simple_Main.PolishingItemSNo,dbo.VRunningLots_Simple_Main.SNO, dbo.VRunningLots_Simple_Main.DT
						,T1.MaxSno, dbo.VRunningLots_Simple_Main.LotType,dbo.VRunningLots_Simple_Main.VRD_EntryID
						,dbo.VRunningLots_Simple_Main.Lot_Available_On_ProcessID
FROM            dbo.VRunningLots_Simple_Main LEFT OUTER JOIN
                             (SELECT        LotNo, MAX(SNO) AS MaxSno
                               FROM            dbo.VRunningLots_Simple_Main AS VRunningLots_Simple_Main_1
                               GROUP BY LotNo) AS T1 ON dbo.VRunningLots_Simple_Main.LotNo = T1.LotNo AND dbo.VRunningLots_Simple_Main.SNO = T1.MaxSno

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRunningLots_Simple]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRunningLots_Simple];
END
GO



CREATE   VIEW [dbo].[VRunningLots_Simple]
AS
SELECT        dbo.VRunningLots_Simple_Main.ItemCode, dbo.VRunningLots_Simple_Main.LotNo, dbo.VRunningLots_Simple_Main.OrderNo, dbo.VRunningLots_Simple_Main.Description, dbo.VRunningLots_Simple_Main.Qty, 
                         dbo.VRunningLots_Simple_Main.ProcessID, dbo.VRunningLots_Simple_Main.VendID, dbo.VRunningLots_Simple_Main.ReWorkLot, dbo.VRunningLots_Simple_Main.R2InDT, dbo.VRunningLots_Simple_Main.PolishingItemSNo, 
                         dbo.VRunningLots_Simple_Main.SNO, dbo.VRunningLots_Simple_Main.DT, T1.LotNo AS Expr1, T1.MaxSno, dbo.VRunningLots_Simple_Main.LotType,dbo.VRunningLots_Simple_Main.VRD_EntryID
                         ,dbo.VRunningLots_Simple_Main.IssEmpID
FROM            dbo.VRunningLots_Simple_Main LEFT OUTER JOIN
                             (SELECT        LotNo, MAX(SNO) AS MaxSno
                               FROM            dbo.VRunningLots_Simple_Main AS VRunningLots_Simple_Main_1
                               GROUP BY LotNo) AS T1 ON dbo.VRunningLots_Simple_Main.LotNo = T1.LotNo AND dbo.VRunningLots_Simple_Main.SNO = T1.MaxSno

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRunningLots_StockOrder]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRunningLots_StockOrder];
END
GO

CREATE   VIEW dbo.VRunningLots_StockOrder
AS
SELECT     OrderNo, ItemCode, RunningQty
FROM        dbo.VRunningLots_Summary
WHERE     (OrderNo IN ('Stock-Order', 'Stock-OrderIss'))

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER OFF;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VVendReceivingList]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VVendReceivingList];
END
GO



CREATE   VIEW [dbo].[VVendReceivingList]
AS
SELECT     dbo.VendReceived.EntryID, dbo.VendReceived.VendID, dbo.VendReceived.DT, dbo.VendReceived.UserID, dbo.VendReceived.ProcessID, dbo.vendRcvdDetail.RecieptID, dbo.vendRcvdDetail.ItemCode, dbo.vendRcvdDetail.RcvdQty, dbo.vendRcvdDetail.Wastage, 
                  dbo.vendRcvdDetail.LostQty, dbo.Makers.VenderName, dbo.Makers.VendID1, CAST(CONVERT(Varchar(10), dbo.VendReceived.DT, 1) AS DATETIME) AS OnlyDT, dbo.vendRcvdDetail.ReqAuth, dbo.VendReceived.Issuance_RefID, dbo.Processes.Description, 
                  dbo.ItemProcesses.Scanning, dbo.vendRcvdDetail.LotNo, dbo.vendRcvdDetail.NextProcessID, dbo.vendRcvdDetail.IssQty, dbo.vendRcvdDetail.EntryID AS VRD_EntryID, dbo.vendRcvdDetail.Opening_RefID, dbo.Items.CatID, dbo.vendRcvdDetail.OrderNo, 
                  dbo.VendIssued.MasterPONo, dbo.vendRcvdDetail.ReWorkLot, FOrderItems.CompItemCode, dbo.vendRcvdDetail.ReWorkQty, dbo.FCustomerOrders.CustCode, dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.Items.GroupID, 
                  dbo.ItemGroups.Description AS ItemGroup, dbo.VendReceived.EmpID, T1.VRD_From_RefID, dbo.VendIssued.DT AS IssDT, dbo.VendIssdDetail.IssQty AS Issuance_IssQty, dbo.VendIssdDetail.RcvdQty AS Issuance_RcvdQty, 
                  dbo.VendReceived_Employees_F(dbo.VendReceived.EntryID) AS Empoloyees, TDL.VRD_RefID, dbo.VendIssdDetail.ReturnDT, dbo.MakerPostedBillsDetail_Receivings.EntryID AS MPB_D_EntryID, dbo.MakerPostedBills.BillNo, dbo.FCustomerOrders.InternalRefNo, 
                  FOrderItems.OrderQty, dbo.VendIssued.Closed, dbo.VItems_Complaints.ItemID AS ComplaintItemID, dbo.Lots_List.Batch_No, dbo.Lots_List.Mill_Certificate_No, dbo.Items.TipSize, dbo.VendReceived.EntryDT, dbo.Employees.Phone1, 
                VHubSupervisors.Hub_Name,VHubSupervisors.Supervisors,
                ISNULL(TDL.DispatchQty, 0) AS DispatchQty,
                ISNULL(TTRF.StoreTransferQty, 0) AS StoreTransferQty,
                CASE WHEN TLC.ClosedLotNo IS NOT NULL THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS IsLotClosed
                ,dbo.VendIssued.BillingDays, dbo.VendIssdDetail.Rate, TLotDispatch.DispatchListNo, 
                 TTRFLD.TTRFD_EntryID
		         ,VRunningLots_Simple.Description AS Lot_Current_Process
		         ,TTotalRcvd.Total_Rcvd_Qty
		         ,TTRFLD.Shelf_RefID,TTRFLD.StoreName,TTRFLD.RackNo,TTRFLD.ShelfNo
FROM        dbo.Employees RIGHT OUTER JOIN
                  dbo.VendReceived INNER JOIN
                  dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID INNER JOIN
                  dbo.Makers ON dbo.VendReceived.VendID = dbo.Makers.VendID INNER JOIN
                  dbo.Processes ON dbo.VendReceived.ProcessID = dbo.Processes.ProcessID INNER JOIN
                  dbo.Items ON dbo.vendRcvdDetail.ItemCode = dbo.Items.ItemID LEFT OUTER JOIN
                  dbo.Lots_List ON dbo.vendRcvdDetail.LotNo = dbo.Lots_List.LotNo LEFT OUTER JOIN
                  dbo.FCustomerOrders ON dbo.vendRcvdDetail.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                  dbo.VItems_Complaints ON dbo.Items.ItemID = dbo.VItems_Complaints.ItemID LEFT OUTER JOIN
                  dbo.MakerPostedBillsDetail_Receivings INNER JOIN
                  dbo.MakerPostedBills ON dbo.MakerPostedBillsDetail_Receivings.MPB_RefID = dbo.MakerPostedBills.EntryID ON dbo.vendRcvdDetail.EntryID = dbo.MakerPostedBillsDetail_Receivings.VRD_RefID LEFT OUTER JOIN
                  dbo.VendIssued INNER JOIN
                  dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID ON dbo.vendRcvdDetail.Issue_RefID = dbo.VendIssdDetail.EntryID LEFT OUTER JOIN
                  dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID ON dbo.Employees.empid = dbo.VendReceived.EmpID LEFT OUTER JOIN
                  dbo.VFOrderItems_Grouped AS FOrderItems ON dbo.vendRcvdDetail.OrderNo = FOrderItems.OrderNo AND dbo.vendRcvdDetail.ItemCode = FOrderItems.CompItemCode LEFT OUTER JOIN
                  dbo.ItemProcesses ON dbo.Processes.ProcessID = dbo.ItemProcesses.ProcessID AND dbo.vendRcvdDetail.ItemCode = dbo.ItemProcesses.ItemID AND dbo.VendReceived.ProcessID = dbo.ItemProcesses.ProcessID LEFT OUTER JOIN
                      (SELECT     VRD_From_RefID
                       FROM        dbo.LotTransferDetails
                       WHERE     (Type = 1)
                       GROUP BY VRD_From_RefID) AS T1 ON dbo.vendRcvdDetail.EntryID = T1.VRD_From_RefID LEFT OUTER JOIN
                      (SELECT     VRD_RefID, SUM(Qty) AS DispatchQty
                       FROM        dbo.DispatchListDetail_VRD
                       GROUP BY VRD_RefID) AS TDL ON dbo.vendRcvdDetail.EntryID = TDL.VRD_RefID 
                       LEFT OUTER JOIN
                       (SELECT LotNo, MAX(DispatchListNo) AS DispatchListNo
                            FROM  dbo.VDispatchListDetail_Adv_Lot
                            GROUP BY LotNo) AS TLotDispatch ON dbo.vendRcvdDetail.LotNo = TLotDispatch.LotNo
                       LEFT OUTER JOIN
                       (SELECT LotNo, MAX(EntryID) AS TTRFD_EntryID,MAX(Shelf_RefID) AS Shelf_RefID,MAX(StoreName) AS StoreName,MAX(ShelfNo) AS ShelfNo,MAX(RackNo) AS RackNo
                        FROM  dbo.VTransferredToReadyFinishLotsDetail
                        GROUP BY LotNo) AS TTRFLD ON dbo.vendRcvdDetail.LotNo = TTRFLD.LotNo
                       LEFT OUTER JOIN
                      (SELECT     VRD_RefID, SUM(Qty) AS StoreTransferQty
                       FROM        dbo.TransferredToReadyFinishLotsDetail
                       GROUP BY VRD_RefID) AS TTRF ON dbo.vendRcvdDetail.EntryID = TTRF.VRD_RefID LEFT OUTER JOIN
                      (SELECT     LotNo AS ClosedLotNo
                       FROM        dbo.Lots_Closed
                       GROUP BY LotNo) AS TLC ON dbo.vendRcvdDetail.LotNo = TLC.ClosedLotNo
                LEFT OUTER JOIN dbo.ItemProcessGroups ON dbo.vendRcvdDetail.ItemCode = dbo.ItemProcessGroups.ItemID 
                LEFT OUTER JOIN dbo.ProcessGroupsProcesses ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroupsProcesses.Group_RefID 
                                                          AND dbo.VendReceived.ProcessID = dbo.ProcessGroupsProcesses.Process_RefID 
                LEFT OUTER JOIN dbo.VHubSupervisors ON dbo.ProcessGroupsProcesses.Group_RefID = dbo.VHubSupervisors.GroupID 
                                                  AND dbo.ProcessGroupsProcesses.Hub_Name = dbo.VHubSupervisors.Hub_Name
                LEFT JOIN VRunningLots_Simple ON dbo.vendRcvdDetail.LotNo=VRunningLots_Simple.LotNo
			    LEFT JOIN (SELECT VendRcvdDetail.Issue_RefID,SUM(VendRcvdDetail.RcvdQty) AS Total_Rcvd_Qty FROM VendRcvdDetail WHERE VendRcvdDetail.Issue_RefID<>0 GROUP BY VendRcvdDetail.Issue_RefID) TTotalRcvd ON VendRcvdDetail.Issue_RefID=TTotalRcvd.Issue_RefID

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER OFF;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VRunningLots_POI]') AND type IN (N'V'))
BEGIN
    DROP VIEW [dbo].[VRunningLots_POI];
END
GO

CREATE   VIEW [dbo].[VRunningLots_POI]
AS
SELECT        TUnion.StartDT, TUnion.LotNo, TUnion.ItemCode, TUnion.Qty, TUnion.Description, TUnion.LastProcessDays, TUnion.OrderNo, TUnion.Remarks, TUnion.DT, TUnion.ProcessID, TUnion.ReWorkLot, 
                         VItems_1.GroupID, FCustomerOrders_1.CustCode, FCustomerOrders_1.Country, TRunning.TLRunning, TRunningOrder.TLRunning_Order, TSF.R1, TSF.R2, TPack.PackStock, VItems_1.ItemName, 
                         VUnshippedOrderItemList_1.OrderNo AS UnShipped_OrderNo, TUnion.EntryType, TUnion.POIRemarks, TUnion.OpenPOQty, TUnion.VendID, TUnion.R1 AS RunningLots_R1, TLOrder.TLUnShippedQty, 
                         TLot_Wastage.Lot_Wastage_Qty, VUnshippedOrderItemList_1.Price, VUnshippedOrderItemList_1.PriceForCost, FOrderItems.Price AS FOI_Price, dbo.ForeignCustomers.Curr, TTray.TrayNo, 
                         TUnion.MasterPONo, FOrderItems.DeliveryDT AS ItemDeliveryDT
						 ,ProcessGroupsProcesses.Hub_Name,TUnion.TrayNo AS Lot_TrayNo
						 ,TDispatch.Qty AS DispatchQty
                         ,FOrderItems.OrderQty
FROM            dbo.ForeignCustomers INNER JOIN
                         dbo.FCustomerOrders AS FCustomerOrders_1 ON dbo.ForeignCustomers.CustCode = FCustomerOrders_1.CustCode AND dbo.ForeignCustomers.Country = FCustomerOrders_1.Country RIGHT OUTER JOIN
                             (SELECT        ISNULL(T1.StartDT, T2.StartDT) AS StartDT, RL.LotNo, RL.ItemCode, RL.Qty, RL.Description, DATEDIFF(day, 
                                                         RL.DT, GETDATE()) AS LastProcessDays, RL.OrderNo, '' AS Remarks, RL.DT, RL.ProcessID, 
                                                         RL.ReWorkLot, 0 AS EntryType, dbo.POI_Remarks.Remarks AS POIRemarks, 0 AS OpenPOQty, RL.VendID, 
                                                         CASE WHEN RL.SNo <= RL.PolishingItemSNo THEN 1 ELSE 0 END AS R1, dbo.VLotsWithMasterPONo.MasterPONo, RL.TrayNo
                               FROM            (SELECT VRL.ItemCode, VRL.LotNo, VRL.OrderNo, VRL.Description, VRL.Qty, VRL.ProcessID, VRL.VendID, VRL.ReWorkLot, VRL.R2InDT, VRL.PolishingItemSNo, VRL.SNO, VRL.DT, VRL.MaxSno, QC.TrayNo
                                                FROM dbo.VRunningLots_Simple VRL WITH (NOLOCK)
                                                LEFT JOIN dbo.VendRcvdDetail_QC_Report QC WITH (NOLOCK) ON VRL.VRD_EntryID = QC.VRD_RefID) RL LEFT OUTER JOIN
                                                         dbo.POI_Remarks ON RL.LotNo = dbo.POI_Remarks.LotNo LEFT OUTER JOIN
                                                             (SELECT        dbo.vendRcvdDetail.LotNo, MIN(dbo.VendReceived.DT) AS StartDT
                                                               FROM            dbo.VendReceived INNER JOIN
                                                                                         dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID
                                                               GROUP BY dbo.vendRcvdDetail.LotNo) AS T1 ON RL.LotNo = T1.LotNo INNER JOIN
                                                         dbo.VItems ON RL.ItemCode = dbo.VItems.ItemID LEFT OUTER JOIN
                                                         dbo.FCustomerOrders ON RL.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                                                             (SELECT        dbo.VendIssdDetail.LotNo, MIN(dbo.VendIssued.DT) AS StartDT
                                                               FROM            dbo.VendIssued INNER JOIN
                                                                                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID
                                                               GROUP BY dbo.VendIssdDetail.LotNo) AS T2 ON RL.LotNo = T2.LotNo LEFT OUTER JOIN
                                                         dbo.VLotsWithMasterPONo ON RL.LotNo = dbo.VLotsWithMasterPONo.LotNo
                               UNION ALL
                               SELECT        dbo.VUnshippedOrderItemList.DT, '' AS LotNo, dbo.VUnshippedOrderItemList.CompItemCode, 
                                                        dbo.VUnshippedOrderItemList.OrderQty - dbo.VUnshippedOrderItemList.ShippedQty - ISNULL(T1_1.RLQty, 0) AS Qty, '' AS Description, 0 AS LastProcessDays, 
                                                        dbo.VUnshippedOrderItemList.OrderNo, dbo.VUnshippedOrderItemList.Remarks, dbo.VUnshippedOrderItemList.DT AS Expr1, 0 AS ProcessID, 0 AS ReWorkLot, 1 AS EntryType, 
                                                        POI_Remarks_1.Remarks AS POIRemarks, TOpenPOs.OpenPOQty, 0 AS VendID, 0 AS R1, TOpenPOs.MasterPONo,'' AS TrayNo
                               FROM            dbo.VUnshippedOrderItemList LEFT OUTER JOIN
                                                        dbo.POI_Remarks AS POI_Remarks_1 ON dbo.VUnshippedOrderItemList.OrderNo = POI_Remarks_1.OrderNo AND 
                                                        dbo.VUnshippedOrderItemList.CompItemCode = POI_Remarks_1.ItemCode LEFT OUTER JOIN
                                                            (SELECT        OrderNo, ItemCode, SUM(Qty) AS OpenPOQty, MAX(MasterPONo) AS MasterPONo
                                                              FROM            dbo.VItemOpenPOs
                                                              GROUP BY OrderNo, ItemCode) AS TOpenPOs ON dbo.VUnshippedOrderItemList.OrderNo = TOpenPOs.OrderNo AND 
                                                        dbo.VUnshippedOrderItemList.CompItemCode = TOpenPOs.ItemCode LEFT OUTER JOIN
                                                            (SELECT        OrderNo, ItemCode, SUM(Qty) AS RLQty
                                                              FROM            (SELECT VRL.ItemCode, VRL.LotNo, VRL.OrderNo, VRL.Qty FROM dbo.VRunningLots_Simple VRL WITH (NOLOCK)) AS VRunningLots_Simple_3
                                                              GROUP BY OrderNo, ItemCode) AS T1_1 ON dbo.VUnshippedOrderItemList.OrderNo = T1_1.OrderNo AND dbo.VUnshippedOrderItemList.CompItemCode = T1_1.ItemCode) AS TUnion INNER JOIN
                         dbo.VItems AS VItems_1 ON TUnion.ItemCode = VItems_1.ItemID INNER JOIN
                         (SELECT OrderNo,CompItemCode,MAX(Price) AS Price,MAX(DeliveryDT) AS DeliveryDT,SUM(FOrderItems.Qty) AS OrderQty FROM FOrderItems GROUP BY OrderNo,CompItemCode) FOrderItems  ON TUnion.ItemCode = FOrderItems.CompItemCode AND TUnion.OrderNo = FOrderItems.OrderNo ON FCustomerOrders_1.OrderNo = TUnion.OrderNo LEFT OUTER JOIN
                             (SELECT        OrderNo, ItemCode, SUM(Qty) AS TLRunning
                               FROM            (SELECT VRL.ItemCode, VRL.LotNo, VRL.OrderNo, VRL.Qty FROM dbo.VRunningLots_Simple VRL WITH (NOLOCK)) AS VRunningLots_Simple_2
                               GROUP BY OrderNo, ItemCode) AS TRunning ON TUnion.OrderNo = TRunning.OrderNo AND TUnion.ItemCode = TRunning.ItemCode LEFT OUTER JOIN
                             (SELECT        OrderNo, ItemCode, SUM(Qty) AS TLRunning_Order
                               FROM            (SELECT VRL.ItemCode, VRL.LotNo, VRL.OrderNo, VRL.Qty FROM dbo.VRunningLots_Simple VRL WITH (NOLOCK)) AS VRunningLots_Simple_1
                               GROUP BY OrderNo, ItemCode) AS TRunningOrder ON TUnion.OrderNo = TRunningOrder.OrderNo AND TUnion.ItemCode = TRunningOrder.ItemCode LEFT OUTER JOIN
                             (SELECT        ItemID, SUM(CASE WHEN R1 = 1 THEN Qty - IssdQty ELSE 0 END) AS R1, SUM(CASE WHEN R1 = 0 THEN Qty - IssdQty ELSE 0 END) AS R2
                               FROM            dbo.VStockOrderOpening
                               WHERE        (Qty > IssdQty)
                               GROUP BY ItemID) AS TSF ON TUnion.ItemCode = TSF.ItemID LEFT OUTER JOIN
                             (SELECT        ItemID, SUM(NetQty) AS PackStock
                               FROM            dbo.VItemsWithShelfWiseStock
                               GROUP BY ItemID) AS TPack ON TUnion.ItemCode = TPack.ItemID LEFT OUTER JOIN
                         dbo.VUnshippedOrderItemList AS VUnshippedOrderItemList_1 ON TUnion.ItemCode = VUnshippedOrderItemList_1.CompItemCode AND TUnion.OrderNo = VUnshippedOrderItemList_1.OrderNo LEFT OUTER JOIN
                             (SELECT        CompItemCode, SUM(OrderQty - ShippedQty) AS TLUnShippedQty
                               FROM            dbo.VUnshippedOrderItemList AS VUnshippedOrderItemList_2
                               GROUP BY CompItemCode) AS TLOrder ON VItems_1.ItemID = TLOrder.CompItemCode LEFT OUTER JOIN
                             (SELECT        LotNo, SUM(Wastage) AS Lot_Wastage_Qty
                               FROM            dbo.vendRcvdDetail AS vendRcvdDetail_1
                               GROUP BY LotNo) AS TLot_Wastage ON TUnion.LotNo = TLot_Wastage.LotNo LEFT OUTER JOIN
                             (SELECT        vendRcvdDetail_2.LotNo, dbo.VendRcvdDetail_QC_Report.TrayNo
                               FROM            dbo.VendRcvdDetail_QC_Report INNER JOIN
                                                         dbo.vendRcvdDetail AS vendRcvdDetail_2 ON vendRcvdDetail_2.EntryID = dbo.VendRcvdDetail_QC_Report.VRD_RefID) AS TTray ON TUnion.LotNo = TTray.LotNo
				LEFT JOIN ItemProcessGroups ON TUnion.ItemCode=ItemProcessGroups.ItemID
				LEFT JOIN ProcessGroupsProcesses ON ItemProcessGroups.PG_RefID=ProcessGroupsProcesses.Group_RefID AND TUnion.ProcessID=ProcessGroupsProcesses.Process_RefID
				LEFT JOIN (SELECT OrderNo,ItemCode,SUM(InnerQty) AS Qty FROM VDispatchListDetail_Ex_Orderwise
								GROUP BY OrderNo,ItemCode) TDispatch ON TUnion.OrderNo=TDispatch.OrderNo AND TUnion.ItemCode=TDispatch.ItemCode

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[EmpGatePassSheet_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[EmpGatePassSheet_SP];
END
GO

CREATE   PROCEDURE EmpGatePassSheet_SP @DTFrom DateTime,@DTTo DateTime
AS
	
BEGIN
	
	SELECT VEmp.EmpID,Name,Rel,FName,StartingSalary,DeptName,DeptID,Designation,Active,Tab2.TotalGPHrs FROM VEmp
		INNER JOIN (SELECT EmpID,SUM(GPHrs) AS TotalGPHrs FROM EmpGPHrs WHERE DT BETWEEN @DTFrom AND @DTTo GROUP BY EmpID) Tab2
	 ON VEmp.EmpID=Tab2.EmpID
	 ORDER BY VEmp.EmpID
END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Financial_Budgeting_Report_Glance]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[Financial_Budgeting_Report_Glance];
END
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE   PROCEDURE [dbo].[Financial_Budgeting_Report_Glance] (@FinancialYear VARCHAR(50),@DTFrom AS DATETIME,@DTTo AS DATETIME)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
CREATE TABLE #Accounts(AccNo VARCHAR(50),AccTitle VARCHAR(255),Proposed BIGINT,Revised BIGINT,AccType BIT)
INSERT INTO #Accounts(AccNo,AccTitle,Proposed,Revised,AccType)
	SELECT Accounts.AccNo,Accounts.AccTitle
			,ISNULL(Month_July,0)+ISNULL(Month_August,0)+ISNULL(Month_September,0)+ISNULL(Month_October,0)+ISNULL(Month_November,0)+ISNULL(Month_December,0)+ISNULL(Month_January,0)+ISNULL(Month_February,0)+ISNULL(Month_March,0)+ISNULL(Month_April,0)+ISNULL(Month_May,0)+ISNULL(Month_June,0) 
			,ISNULL(Month_July_R,0)+ISNULL(Month_August_R,0)+ISNULL(Month_September_R,0)+ISNULL(Month_October_R,0)+ISNULL(Month_November_R,0)+ISNULL(Month_December_R,0)+ISNULL(Month_January_R,0)+ISNULL(Month_February_R,0)+ISNULL(Month_March_R,0)+ISNULL(Month_April_R,0)+ISNULL(Month_May_R,0)+ISNULL(Month_June_R,0)
			,0
			FROM Accounts
		INNER JOIN FinancialBudgeting T1 ON Accounts.AccNo=T1.AccNo
		WHERE BudgetAccount_Glance=1
		AND T1.FinancialYear=@FinancialYear
	UNION ALL
	SELECT Accounts.AccNo,Accounts.AccTitle
			,Proposed
			,Revised
			,1
			FROM Accounts
		INNER JOIN FinancialBudgeting_Expenses T1 ON Accounts.AccNo=T1.AccNo
		WHERE BudgetAccount_Glance=1
		AND T1.FinancialYear=@FinancialYear
	UNION ALL
	SELECT Code,Head
		,SUM(ISNULL(Month_July,0)+ISNULL(Month_August,0)+ISNULL(Month_September,0)+ISNULL(Month_October,0)+ISNULL(Month_November,0)+ISNULL(Month_December,0)+ISNULL(Month_January,0)+ISNULL(Month_February,0)+ISNULL(Month_March,0)+ISNULL(Month_April,0)+ISNULL(Month_May,0)+ISNULL(Month_June,0))
		,SUM(ISNULL(Month_July_R,0)+ISNULL(Month_August_R,0)+ISNULL(Month_September_R,0)+ISNULL(Month_October_R,0)+ISNULL(Month_November_R,0)+ISNULL(Month_December_R,0)+ISNULL(Month_January_R,0)+ISNULL(Month_February_R,0)+ISNULL(Month_March_R,0)+ISNULL(Month_April_R,0)+ISNULL(Month_May_R,0)+ISNULL(Month_June_R,0))
		,0
		FROM Heads
		INNER JOIN FinancialBudgeting T1 ON LEFT(T1.AccNo,2)=Heads.Code
		WHERE BudgetAccount_Glance=1
		AND T1.FinancialYear=@FinancialYear
		GROUP BY Heads.Code,Heads.Head
	UNION ALL
	SELECT Code,Head
		,SUM(Proposed)
		,SUM(Revised)
		,1
		FROM Heads
		INNER JOIN FinancialBudgeting_Expenses T1 ON LEFT(T1.AccNo,2)=Heads.Code
		WHERE BudgetAccount_Glance=1
		AND T1.FinancialYear=@FinancialYear
		GROUP BY Heads.Code,Heads.Head


SELECT AccNo,#Accounts.AccTitle,Proposed,Revised
		,dbo.GetTransactionValue(AccNo,@DTFrom,@DTTo) AS Actual,AccType AS Expense
	FROM #Accounts	
	WHERE AccType=0
	
UNION ALL

SELECT AccNo,#Accounts.AccTitle,Proposed,Revised,dbo.GetTransactionValue(AccNo,@DTFrom,@DTTo) AS Actual,AccType AS Expense
	 FROM #Accounts
	WHERE AccType=1
	
END

DROP TABLE #Accounts

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FinancialBudgeting_Expenses_Report_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[FinancialBudgeting_Expenses_Report_SP];
END
GO

CREATE   PROCEDURE [dbo].[FinancialBudgeting_Expenses_Report_SP] 
	(@FinancialYear VARCHAR(50),@DTFrom AS DATETIME,@DTTo AS DATETIME,@QuarterNo AS INT,@Type AS VARCHAR(50)='0',@DateRangeWise AS BIT=0,@NoOfMonths INT=0) AS
BEGIN
	SET NOCOUNT ON;

	SET @DTFrom = CAST(@DTFrom AS DATE);
	SET @DTTo = CAST(@DTTo AS DATE);

	DECLARE @DTFrom1 AS DATETIME, @DTTo1 AS DATETIME;
	DECLARE @DTFrom2 AS DATETIME, @DTTo2 AS DATETIME;
	DECLARE @DTFrom3 AS DATETIME, @DTTo3 AS DATETIME;
	DECLARE @DTFrom4 AS DATETIME, @DTTo4 AS DATETIME;

	DECLARE @DTFrom_ToUse DATETIME, @DTTo_ToUse DATETIME;
	DECLARE @QuarterName VARCHAR(50);

	SET @DTFrom1 = @DTFrom;
	SET @DTFrom2 = DATEADD(month, 3, @DTFrom);
	SET @DTFrom3 = DATEADD(month, 6, @DTFrom);
	SET @DTFrom4 = DATEADD(month, 9, @DTFrom);

	SET @DTTo1 = DATEADD(day, -1, @DTFrom2);
	SET @DTTo2 = DATEADD(day, -1, @DTFrom3);
	SET @DTTo3 = DATEADD(day, -1, DATEADD(month, 1, @DTFrom4));
	SET @DTTo4 = @DTTo;

	IF @DateRangeWise = 1
	BEGIN
		SET @DTFrom_ToUse = @DTFrom;
		SET @DTTo_ToUse = @DTTo;
		SET @QuarterName = LEFT(DATENAME(MONTH, @DTFrom), 3) + '-' + LEFT(DATENAME(MONTH, @DTTo), 3);
	END
	ELSE IF @QuarterNo = 1
	BEGIN
		SET @DTFrom_ToUse = @DTFrom1;
		SET @DTTo_ToUse = @DTTo1;
		SET @QuarterName = 'Jul-Sep';
	END
	ELSE IF @QuarterNo = 2
	BEGIN
		SET @DTFrom_ToUse = @DTFrom;
		SET @DTTo_ToUse = @DTTo2;
		SET @QuarterName = 'Jul-Dec';
	END
	ELSE IF @QuarterNo = 3
	BEGIN
		SET @DTFrom_ToUse = @DTFrom;
		SET @DTTo_ToUse = @DTTo3;
		SET @QuarterName = 'Jul-Mar';
	END
	ELSE IF @QuarterNo = 4
	BEGIN
		SET @DTFrom_ToUse = @DTFrom;
		SET @DTTo_ToUse = @DTTo4;
		SET @QuarterName = 'Jul-Jun';
	END;

	SELECT 
		ISNULL(FinancialBudgeting_Expenses.EntryID, 0) AS EntryID,
		@FinancialYear AS FinancialYear,
		VAccounts.AccNo,
		VAccounts.AccTitle,
		ISNULL(FinancialBudgeting_Expenses.Proposed, 0) AS YearlyBudget,
		ISNULL(FinancialBudgeting_Expenses.Revised, 0) AS YearlyBudget_Revision,
		CASE WHEN @DateRangeWise = 1 THEN
			(ISNULL(FinancialBudgeting_Expenses.Proposed, 0) * @NoOfMonths) / 12
		ELSE
			CASE @QuarterNo
				WHEN 1 THEN (ISNULL(FinancialBudgeting_Expenses.Proposed, 0) * 3) / 12
				WHEN 2 THEN (ISNULL(FinancialBudgeting_Expenses.Proposed, 0) * 6) / 12
				WHEN 3 THEN (ISNULL(FinancialBudgeting_Expenses.Proposed, 0) * 9) / 12
				WHEN 4 THEN (ISNULL(FinancialBudgeting_Expenses.Proposed, 0) * 12) / 12
				ELSE 0
			END
		END AS Quarter_Budget,
		CASE WHEN @DateRangeWise = 1 THEN
			(ISNULL(FinancialBudgeting_Expenses.Revised, 0) * @NoOfMonths) / 12
		ELSE
			CASE @QuarterNo
				WHEN 1 THEN (ISNULL(FinancialBudgeting_Expenses.Revised, 0) * 3) / 12
				WHEN 2 THEN (ISNULL(FinancialBudgeting_Expenses.Revised, 0) * 6) / 12
				WHEN 3 THEN (ISNULL(FinancialBudgeting_Expenses.Revised, 0) * 9) / 12
				WHEN 4 THEN (ISNULL(FinancialBudgeting_Expenses.Revised, 0) * 12) / 12
				ELSE 0
			END
		END AS Quarter_Budget_Revision,
		ISNULL(dbo.GetTransactionValue(VAccounts.AccNo, @DTFrom_ToUse, @DTTo_ToUse), 0) AS Quarter_Actual,
		@QuarterName AS QuarterName
	FROM VAccounts
	LEFT JOIN FinancialBudgeting_Expenses ON VAccounts.AccNo = FinancialBudgeting_Expenses.AccNo AND FinancialBudgeting_Expenses.FinancialYear = @FinancialYear
	WHERE VAccounts.BudgetAccount = 1
	AND VAccounts.Active = 1
	AND (TRY_CAST(VAccounts.Type AS INT) >= 41 OR LEFT(VAccounts.AccNo, 6) = '12-001')
	AND (VAccounts.Type = @Type OR @Type = '0')
	ORDER BY VAccounts.AccNo;
END;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FinancialBudgeting_Expenses_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[FinancialBudgeting_Expenses_SP];
END
GO

CREATE   PROCEDURE [dbo].[FinancialBudgeting_Expenses_SP] (@FinancialYear VARCHAR(50),@DTFrom As DateTime,@DTTo As DateTime,@Type AS VARCHAR(50)='0') AS

SET @DTFrom=Convert(varchar(8),@DTFrom,1);SET @DTTo=Convert(varchar(8),@DTTo,1)

/*DECLARE @DTFrom1 AS DATETIME,@DTTo1 AS DATETIME
DECLARE @DTFrom2 AS DATETIME,@DTTo2 AS DATETIME
DECLARE @DTFrom3 AS DATETIME,@DTTo3 AS DATETIME
DECLARE @DTFrom4 AS DATETIME,@DTTo4 AS DATETIME
DECLARE @DTFrom5 AS DATETIME,@DTTo5 AS DATETIME
DECLARE @DTFrom6 AS DATETIME,@DTTo6 AS DATETIME
DECLARE @DTFrom7 AS DATETIME,@DTTo7 AS DATETIME
DECLARE @DTFrom8 AS DATETIME,@DTTo8 AS DATETIME
DECLARE @DTFrom9 AS DATETIME,@DTTo9 AS DATETIME
DECLARE @DTFrom10 AS DATETIME,@DTTo10 AS DATETIME
DECLARE @DTFrom11 AS DATETIME,@DTTo11 AS DATETIME
DECLARE @DTFrom12 AS DATETIME,@DTTo12 AS DATETIME

SET @DTFrom1=@DTFrom;SET @DTFrom2=DateAdd(m,1,@DTFrom);SET @DTFrom3=DateAdd(m,2,@DTFrom);SET @DTFrom4=DateAdd(m,3,@DTFrom)
SET @DTFrom5=DateAdd(m,4,@DTFrom);SET @DTFrom6=DateAdd(m,5,@DTFrom);SET @DTFrom7=DateAdd(m,6,@DTFrom);SET @DTFrom8=DateAdd(m,7,@DTFrom)
SET @DTFrom9=DateAdd(m,8,@DTFrom);SET @DTFrom10=DateAdd(m,9,@DTFrom);SET @DTFrom11=DateAdd(m,10,@DTFrom);SET @DTFrom12=DateAdd(m,11,@DTFrom)

SET @DTTo1=DateAdd(d,-1,DateAdd(m,1,@DTFrom1));SET @DTTo2=DateAdd(d,-1,DateAdd(m,1,@DTFrom2));SET @DTTo3=DateAdd(d,-1,DateAdd(m,1,@DTFrom3))
SET @DTTo4=DateAdd(d,-1,DateAdd(m,1,@DTFrom4));SET @DTTo5=DateAdd(d,-1,DateAdd(m,1,@DTFrom5));SET @DTTo6=DateAdd(d,-1,DateAdd(m,1,@DTFrom6))
SET @DTTo7=DateAdd(d,-1,DateAdd(m,1,@DTFrom7));SET @DTTo8=DateAdd(d,-1,DateAdd(m,1,@DTFrom8));SET @DTTo9=DateAdd(d,-1,DateAdd(m,1,@DTFrom9))
SET @DTTo10=DateAdd(d,-1,DateAdd(m,1,@DTFrom10));SET @DTTo11=DateAdd(d,-1,DateAdd(m,1,@DTFrom11));SET @DTTo12=DateAdd(d,-1,DateAdd(m,1,@DTFrom12))*/


SELECT EntryID,FinancialBudgeting_Expenses.FinancialYear,FinancialBudgeting_Expenses.AccNo,VAccounts.AccTitle,FinancialBudgeting_Expenses.Proposed,FinancialBudgeting_Expenses.Revised 
	FROM FinancialBudgeting_Expenses
	/*LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_July FROM Vouchers WHERE VDate BETWEEN @DTFrom1 AND @DTTo1 GROUP BY AccNo) TJul ON FinancialBudgeting.AccNo=TJul.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_August FROM Vouchers WHERE VDate BETWEEN @DTFrom2 AND @DTTo2 GROUP BY AccNo) TAug ON FinancialBudgeting.AccNo=TAug.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_September FROM Vouchers WHERE VDate BETWEEN @DTFrom3 AND @DTTo3 GROUP BY AccNo) TSep ON FinancialBudgeting.AccNo=TSep.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_October FROM Vouchers WHERE VDate BETWEEN @DTFrom4 AND @DTTo4 GROUP BY AccNo) TOct ON FinancialBudgeting.AccNo=TOct.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_November FROM Vouchers WHERE VDate BETWEEN @DTFrom5 AND @DTTo5 GROUP BY AccNo) TNov ON FinancialBudgeting.AccNo=TNov.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_December FROM Vouchers WHERE VDate BETWEEN @DTFrom6 AND @DTTo6 GROUP BY AccNo) TDec ON FinancialBudgeting.AccNo=TDec.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_January FROM Vouchers WHERE VDate BETWEEN @DTFrom7 AND @DTTo7 GROUP BY AccNo) TJan ON FinancialBudgeting.AccNo=TJan.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_February FROM Vouchers WHERE VDate BETWEEN @DTFrom8 AND @DTTo8 GROUP BY AccNo) TFeb ON FinancialBudgeting.AccNo=TFeb.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_March FROM Vouchers WHERE VDate BETWEEN @DTFrom9 AND @DTTo9 GROUP BY AccNo) TMar ON FinancialBudgeting.AccNo=TMar.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_April FROM Vouchers WHERE VDate BETWEEN @DTFrom10 AND @DTTo10 GROUP BY AccNo) TApr ON FinancialBudgeting.AccNo=TApr.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_May FROM Vouchers WHERE VDate BETWEEN @DTFrom11 AND @DTTo11 GROUP BY AccNo) TMay ON FinancialBudgeting.AccNo=TMay.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_Jun FROM Vouchers WHERE VDate BETWEEN @DTFrom12 AND @DTTo12 GROUP BY AccNo) TJun ON FinancialBudgeting.AccNo=TJun.AccNo*/
	INNER JOIN VAccounts ON FinancialBudgeting_Expenses.AccNo=VAccounts.AccNo
	AND VAccounts.BudgetAccount=1
	WHERE FinancialBudgeting_Expenses.FinancialYear=@FinancialYear
	AND (VACcounts.Type=@Type OR @Type='0')

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FinancialBudgeting_Monthwise_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[FinancialBudgeting_Monthwise_SP];
END
GO

CREATE   PROCEDURE [dbo].[FinancialBudgeting_Monthwise_SP] (@FinancialYear VARCHAR(50),@DTFrom As DateTime,@DTTo As DateTime,@Type VARCHAR(50)='0') AS

SET @DTFrom=Convert(varchar(8),@DTFrom,1);SET @DTTo=Convert(varchar(8),@DTTo,1)

DECLARE @DTFrom1 AS DATETIME,@DTTo1 AS DATETIME
DECLARE @DTFrom2 AS DATETIME,@DTTo2 AS DATETIME
DECLARE @DTFrom3 AS DATETIME,@DTTo3 AS DATETIME
DECLARE @DTFrom4 AS DATETIME,@DTTo4 AS DATETIME
DECLARE @DTFrom5 AS DATETIME,@DTTo5 AS DATETIME
DECLARE @DTFrom6 AS DATETIME,@DTTo6 AS DATETIME
DECLARE @DTFrom7 AS DATETIME,@DTTo7 AS DATETIME
DECLARE @DTFrom8 AS DATETIME,@DTTo8 AS DATETIME
DECLARE @DTFrom9 AS DATETIME,@DTTo9 AS DATETIME
DECLARE @DTFrom10 AS DATETIME,@DTTo10 AS DATETIME
DECLARE @DTFrom11 AS DATETIME,@DTTo11 AS DATETIME
DECLARE @DTFrom12 AS DATETIME,@DTTo12 AS DATETIME

SET @DTFrom1=@DTFrom;SET @DTFrom2=DateAdd(m,1,@DTFrom);SET @DTFrom3=DateAdd(m,2,@DTFrom);SET @DTFrom4=DateAdd(m,3,@DTFrom)
SET @DTFrom5=DateAdd(m,4,@DTFrom);SET @DTFrom6=DateAdd(m,5,@DTFrom);SET @DTFrom7=DateAdd(m,6,@DTFrom);SET @DTFrom8=DateAdd(m,7,@DTFrom)
SET @DTFrom9=DateAdd(m,8,@DTFrom);SET @DTFrom10=DateAdd(m,9,@DTFrom);SET @DTFrom11=DateAdd(m,10,@DTFrom);SET @DTFrom12=DateAdd(m,11,@DTFrom)

SET @DTTo1=DateAdd(d,-1,DateAdd(m,1,@DTFrom1));SET @DTTo2=DateAdd(d,-1,DateAdd(m,1,@DTFrom2));SET @DTTo3=DateAdd(d,-1,DateAdd(m,1,@DTFrom3))
SET @DTTo4=DateAdd(d,-1,DateAdd(m,1,@DTFrom4));SET @DTTo5=DateAdd(d,-1,DateAdd(m,1,@DTFrom5));SET @DTTo6=DateAdd(d,-1,DateAdd(m,1,@DTFrom6))
SET @DTTo7=DateAdd(d,-1,DateAdd(m,1,@DTFrom7));SET @DTTo8=DateAdd(d,-1,DateAdd(m,1,@DTFrom8));SET @DTTo9=DateAdd(d,-1,DateAdd(m,1,@DTFrom9))
SET @DTTo10=DateAdd(d,-1,DateAdd(m,1,@DTFrom10));SET @DTTo11=DateAdd(d,-1,DateAdd(m,1,@DTFrom11));SET @DTTo12=DateAdd(d,-1,DateAdd(m,1,@DTFrom12))


SELECT EntryID,FinancialYear,FinancialBudgeting.AccNo,VAccounts.AccTitle,Month_July,Month_August,Month_September,Month_October,Month_November,Month_December,Month_January,Month_February,Month_March,Month_April,Month_May,Month_June
	,Month_July_R,Month_August_R,Month_September_R,Month_October_R,Month_November_R,Month_December_R,Month_January_R,Month_February_R,Month_March_R,Month_April_R,Month_May_R,Month_June_R
	,Month_A_July,Month_A_August,Month_A_September,Month_A_October,Month_A_November,Month_A_December,Month_A_January,Month_A_February,Month_A_March,Month_A_April,Month_A_May,Month_A_Jun FROM FinancialBudgeting
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_July FROM Vouchers WHERE VDate BETWEEN @DTFrom1 AND @DTTo1 GROUP BY AccNo) TJul ON FinancialBudgeting.AccNo=TJul.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_August FROM Vouchers WHERE VDate BETWEEN @DTFrom2 AND @DTTo2 GROUP BY AccNo) TAug ON FinancialBudgeting.AccNo=TAug.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_September FROM Vouchers WHERE VDate BETWEEN @DTFrom3 AND @DTTo3 GROUP BY AccNo) TSep ON FinancialBudgeting.AccNo=TSep.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_October FROM Vouchers WHERE VDate BETWEEN @DTFrom4 AND @DTTo4 GROUP BY AccNo) TOct ON FinancialBudgeting.AccNo=TOct.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_November FROM Vouchers WHERE VDate BETWEEN @DTFrom5 AND @DTTo5 GROUP BY AccNo) TNov ON FinancialBudgeting.AccNo=TNov.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_December FROM Vouchers WHERE VDate BETWEEN @DTFrom6 AND @DTTo6 GROUP BY AccNo) TDec ON FinancialBudgeting.AccNo=TDec.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_January FROM Vouchers WHERE VDate BETWEEN @DTFrom7 AND @DTTo7 GROUP BY AccNo) TJan ON FinancialBudgeting.AccNo=TJan.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_February FROM Vouchers WHERE VDate BETWEEN @DTFrom8 AND @DTTo8 GROUP BY AccNo) TFeb ON FinancialBudgeting.AccNo=TFeb.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_March FROM Vouchers WHERE VDate BETWEEN @DTFrom9 AND @DTTo9 GROUP BY AccNo) TMar ON FinancialBudgeting.AccNo=TMar.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_April FROM Vouchers WHERE VDate BETWEEN @DTFrom10 AND @DTTo10 GROUP BY AccNo) TApr ON FinancialBudgeting.AccNo=TApr.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_May FROM Vouchers WHERE VDate BETWEEN @DTFrom11 AND @DTTo11 GROUP BY AccNo) TMay ON FinancialBudgeting.AccNo=TMay.AccNo
	LEFT OUTER JOIN (SELECT AccNo,SUM(Debit-Credit) AS Month_A_Jun FROM Vouchers WHERE VDate BETWEEN @DTFrom12 AND @DTTo12 GROUP BY AccNo) TJun ON FinancialBudgeting.AccNo=TJun.AccNo
	INNER JOIN VAccounts ON FinancialBudgeting.AccNo=VAccounts.AccNo
	AND VAccounts.BudgetAccount=1
	WHERE FinancialYear=@FinancialYear
	AND (VAccounts.Type=@Type OR @Type='0')

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FinancialBudgeting_Report_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[FinancialBudgeting_Report_SP];
END
GO

CREATE   PROCEDURE [dbo].[FinancialBudgeting_Report_SP]
	(@FinancialYear VARCHAR(50),@DTFrom AS DATETIME,@DTTo AS DATETIME,
	@QuarterNo AS INT=0,@Type AS VARCHAR(50)='0',@DateRangeWise AS BIT=0,
	@Proposed INT=0,@Revised INT=0) AS
BEGIN
	SET NOCOUNT ON;

	SET @DTFrom = CAST(@DTFrom AS DATE);
	SET @DTTo = CAST(@DTTo AS DATE);

	DECLARE @DTFrom1 AS DATETIME, @DTTo1 AS DATETIME;
	DECLARE @DTFrom2 AS DATETIME, @DTTo2 AS DATETIME;
	DECLARE @DTFrom3 AS DATETIME, @DTTo3 AS DATETIME;
	DECLARE @DTFrom4 AS DATETIME, @DTTo4 AS DATETIME;

	DECLARE @DTFrom_ToUse DATETIME, @DTTo_ToUse DATETIME;
	DECLARE @QuarterName VARCHAR(50);

	SET @DTFrom1 = @DTFrom;
	SET @DTFrom2 = DATEADD(month, 3, @DTFrom);
	SET @DTFrom3 = DATEADD(month, 6, @DTFrom);
	SET @DTFrom4 = DATEADD(month, 9, @DTFrom);

	SET @DTTo1 = DATEADD(day, -1, @DTFrom2);
	SET @DTTo2 = DATEADD(day, -1, @DTFrom3);
	SET @DTTo3 = DATEADD(day, -1, DATEADD(month, 1, @DTFrom4));
	SET @DTTo4 = @DTTo;

	CREATE TABLE #AccNos(AccNo VARCHAR(50) COLLATE database_default, Proposed INT, Revised INT);

	IF @DateRangeWise = 1
	BEGIN
		DECLARE @DTTemp DATETIME;
		SET @DTTemp = @DTFrom;
		DECLARE @Months VARCHAR(1000) = '', @Months_R VARCHAR(1000) = '', @Query VARCHAR(2000);

		WHILE @DTTemp <= @DTTo
		BEGIN
			SET @Months = @Months + 'ISNULL(Month_' + DATENAME(MONTH, @DTTemp) + ',0)+';
			SET @Months_R = @Months_R + 'ISNULL(Month_' + DATENAME(MONTH, @DTTemp) + '_R,0)+';
			SET @DTTemp = DATEADD(month, 1, @DTTemp);
		END;

		IF @Months <> ''
			SET @Months = LEFT(@Months, LEN(@Months) - 1);

		IF @Months_R <> ''
			SET @Months_R = LEFT(@Months_R, LEN(@Months_R) - 1);

		IF @Months <> ''
		BEGIN
			SET @Query = 'INSERT INTO #AccNos(AccNo, Proposed) SELECT AccNo, ' + @Months + ' FROM FinancialBudgeting WHERE FinancialYear=''' + REPLACE(@FinancialYear, '''', '''''') + '''';
			EXEC(@Query);

			SET @Query = 'UPDATE T1 SET T1.Revised=' + @Months_R + ' FROM #AccNos T1 INNER JOIN FinancialBudgeting ON T1.AccNo=FinancialBudgeting.AccNo WHERE FinancialYear=''' + REPLACE(@FinancialYear, '''', '''''') + '''';
			EXEC(@Query);
		END;
	END;

	IF @DateRangeWise = 1
	BEGIN
		SET @DTFrom_ToUse = @DTFrom;
		SET @DTTo_ToUse = @DTTo;
		SET @QuarterName = LEFT(DATENAME(MONTH, @DTFrom), 3) + '-' + LEFT(DATENAME(MONTH, @DTTo), 3);
	END
	ELSE IF @QuarterNo = 1
	BEGIN
		SET @DTFrom_ToUse = @DTFrom1;
		SET @DTTo_ToUse = @DTTo1;
		SET @QuarterName = 'Jul-Sep';
	END
	ELSE IF @QuarterNo = 2
	BEGIN
		SET @DTFrom_ToUse = @DTFrom;
		SET @DTTo_ToUse = @DTTo2;
		SET @QuarterName = 'Jul-Dec';
	END
	ELSE IF @QuarterNo = 3
	BEGIN
		SET @DTFrom_ToUse = @DTFrom;
		SET @DTTo_ToUse = @DTTo3;
		SET @QuarterName = 'Jul-Mar';
	END
	ELSE IF @QuarterNo = 4
	BEGIN
		SET @DTFrom_ToUse = @DTFrom;
		SET @DTTo_ToUse = @DTTo4;
		SET @QuarterName = 'Jul-Jun';
	END;

	SELECT 
		ISNULL(FinancialBudgeting.EntryID, 0) AS EntryID,
		@FinancialYear AS FinancialYear,
		VAccounts.AccNo,
		VAccounts.AccTitle,
		ISNULL(Month_July,0)+ISNULL(Month_August,0)+ISNULL(Month_September,0)+ISNULL(Month_October,0)+ISNULL(Month_November,0)+ISNULL(Month_December,0)+ISNULL(Month_January,0)+ISNULL(Month_February,0)+ISNULL(Month_March,0)+ISNULL(Month_April,0)+ISNULL(Month_May,0)+ISNULL(Month_June,0) AS YearlyBudget,
		ISNULL(Month_July_R,0)+ISNULL(Month_August_R,0)+ISNULL(Month_September_R,0)+ISNULL(Month_October_R,0)+ISNULL(Month_November_R,0)+ISNULL(Month_December_R,0)+ISNULL(Month_January_R,0)+ISNULL(Month_February_R,0)+ISNULL(Month_March_R,0)+ISNULL(Month_April_R,0)+ISNULL(Month_May_R,0)+ISNULL(Month_June_R,0) AS YearlyBudget_Revision,
		CASE WHEN @DateRangeWise = 1 THEN
			ISNULL(T1.Proposed, 0)
		ELSE
			CASE @QuarterNo
				WHEN 1 THEN ISNULL(Month_July,0)+ISNULL(Month_August,0)+ISNULL(Month_September,0)
				WHEN 2 THEN ISNULL(Month_July,0)+ISNULL(Month_August,0)+ISNULL(Month_September,0)+ISNULL(Month_October,0)+ISNULL(Month_November,0)+ISNULL(Month_December,0)
				WHEN 3 THEN ISNULL(Month_July,0)+ISNULL(Month_August,0)+ISNULL(Month_September,0)+ISNULL(Month_October,0)+ISNULL(Month_November,0)+ISNULL(Month_December,0)+ISNULL(Month_January,0)+ISNULL(Month_February,0)+ISNULL(Month_March,0)
				WHEN 4 THEN ISNULL(Month_July,0)+ISNULL(Month_August,0)+ISNULL(Month_September,0)+ISNULL(Month_October,0)+ISNULL(Month_November,0)+ISNULL(Month_December,0)+ISNULL(Month_January,0)+ISNULL(Month_February,0)+ISNULL(Month_March,0)+ISNULL(Month_April,0)+ISNULL(Month_May,0)+ISNULL(Month_June,0)
				ELSE 0
			END
		END AS Quarter_Budget,
		CASE WHEN @DateRangeWise = 1 THEN
			ISNULL(T1.Revised, 0)
		ELSE
			CASE @QuarterNo
				WHEN 1 THEN ISNULL(Month_July_R,0)+ISNULL(Month_August_R,0)+ISNULL(Month_September_R,0)
				WHEN 2 THEN ISNULL(Month_July_R,0)+ISNULL(Month_August_R,0)+ISNULL(Month_September_R,0)+ISNULL(Month_October_R,0)+ISNULL(Month_November_R,0)+ISNULL(Month_December_R,0)
				WHEN 3 THEN ISNULL(Month_July_R,0)+ISNULL(Month_August_R,0)+ISNULL(Month_September_R,0)+ISNULL(Month_October_R,0)+ISNULL(Month_November_R,0)+ISNULL(Month_December_R,0)+ISNULL(Month_January_R,0)+ISNULL(Month_February_R,0)+ISNULL(Month_March_R,0)
				WHEN 4 THEN ISNULL(Month_July_R,0)+ISNULL(Month_August_R,0)+ISNULL(Month_September_R,0)+ISNULL(Month_October_R,0)+ISNULL(Month_November_R,0)+ISNULL(Month_December_R,0)+ISNULL(Month_January_R,0)+ISNULL(Month_February_R,0)+ISNULL(Month_March_R,0)+ISNULL(Month_April_R,0)+ISNULL(Month_May_R,0)+ISNULL(Month_June_R,0)
				ELSE 0
			END
		END AS Quarter_Budget_Revision,
		ISNULL(dbo.GetTransactionValue(VAccounts.AccNo, @DTFrom_ToUse, @DTTo_ToUse), 0) AS Quarter_Actual,
		ISNULL(@QuarterName, '') AS QuarterName
	FROM VAccounts
	LEFT JOIN FinancialBudgeting ON VAccounts.AccNo = FinancialBudgeting.AccNo AND FinancialBudgeting.FinancialYear = @FinancialYear
	LEFT JOIN #AccNos T1 ON VAccounts.AccNo = T1.AccNo
	WHERE VAccounts.BudgetAccount = 1
	AND VAccounts.Active = 1
	AND VAccounts.Type IN ('31', '32', '33')
	AND (VAccounts.Type = @Type OR @Type = '0')
	ORDER BY VAccounts.AccNo;

	DROP TABLE #AccNos;
END;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[InsertChequeDetails_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[InsertChequeDetails_SP];
END
GO
CREATE   PROCEDURE [dbo].[InsertChequeDetails_SP]
    @SNo [bigint] ,
	@CDate [smalldatetime] NULL,
	@AccNo [nvarchar](255) NULL,
	@Description [nvarchar](4000) NULL,
	@Amount [money] NULL,
	@BankID [int] NULL,
	@chequeno [nvarchar](50) NULL,
	@ChqBookNo [int] NULL,
	@chequeType [nvarchar](50) NULL,
	@chequeDate [smalldatetime] NULL,
	@Posted [bit] ,
	@Payment [bit] ,
	@Bounced [bit] ,
	@ChqIsDue [bit] NULL,
	@Payee [varchar](4000) NULL,
	@BankAccNo [varchar](50) NULL,
	@ClearanceDT [datetime] NULL,
	@ChqPrintingDone [bit] NULL
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO [dbo].[Cheque]
           ([SNo]
           ,[CDate]
           ,[AccNo]
           ,[Description]
           ,[Amount]
           ,[BankID]
           ,[chequeno]
           ,[ChqBookNo]
           ,[chequeType]
           ,[chequeDate]
           ,[Posted]
           ,[Payment]
           ,[Bounced]
           ,[ChqIsDue]
           ,[Payee]
           ,[BankAccNo]
           ,[ClearanceDT]
           ,[ChqPrintingDone]
           )                
    OUTPUT INSERTED.SNo  -- ***CRITICAL: Returns the new ID***
    VALUES 
    (
        @SNo,
		@CDate,
		@AccNo,
		@Description,
		@Amount,
		@BankID,
		@chequeno,
		@ChqBookNo,
		@chequeType,
		@chequeDate,
		@Posted,
		@Payment,
		@Bounced,
		@ChqIsDue,
		@Payee,
		@BankAccNo,
		@ClearanceDT,
		@ChqPrintingDone
    );

	--Commented the following line as it's being handled in InsertVoucherHead_SP SP.
	--UPDATE ChqList SET Issued=1 WHERE ChqBookNo=@ChqBookNo AND ChqNo=@chequeno

END


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[InsertChqBookData_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[InsertChqBookData_SP];
END
GO

CREATE   PROCEDURE 
	[dbo].[InsertChqBookData_SP]
    @StartingFrom INT,
    @Chqs INT,
    @ManualNo NVARCHAR(50),
    @ChqBookDetail NVARCHAR(200),
    @AccNo NVARCHAR(50)	
AS
BEGIN

	SET NOCOUNT ON;
	
	INSERT INTO ChqBooks(AccNo,StartingFrom,Chqs,ManualNo,ChqBookDetail)
	VALUES 
	(@AccNo,@StartingFrom,@Chqs,@ManualNo,@ChqBookDetail)

	SELECT MAX(ChqBookNo) FROM ChqBooks;
END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[InsertChqBookDetailData_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[InsertChqBookDetailData_SP];
END
GO

CREATE   PROCEDURE [dbo].[InsertChqBookDetailData_SP] (@StartingFrom VARCHAR(50),@Chqs INT,@ChqBookNo AS INT)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @iSrNo INT,@ChqLen INT,@EndNo INT

    --SET @iSrNo = CAST(@StartingFrom AS INT);
    SET @iSrNo = 0;
    SET @ChqLen = LEN(@StartingFrom);
    SET @EndNo = @StartingFrom + @Chqs - 1;

    --CREATE TABLE #Temp(ChqBookNo VARCHAR(50), ChqNo VARCHAR(50),Issued tinyint)
    WHILE @iSrNo <= (@Chqs-1)
    BEGIN
        INSERT INTO ChqList (ChqBookNo, ChqNo,Issued)
        --INSERT INTO #Temp(ChqBookNo, ChqNo,Issued)
        VALUES (@ChqBookNo,FORMAT(@StartingFrom + @iSrNo, REPLICATE('0', @ChqLen)),0);

        SET @iSrNo = @iSrNo + 1;
    END

    SELECT @EndNo AS LastChequeNo;
    --SELECT * FROM #Temp
    --DROP TABLE #Temp;

END;

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[InsertVoucherHead_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[InsertVoucherHead_SP];
END
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE   PROCEDURE [dbo].[InsertVoucherHead_SP] 
				(

					@VchrNo [varchar](50) ,
					@UserName [varchar](50) NULL,
					@MachineName [varchar](50) NULL,
					@DT [datetime] NULL,
					@Notes [varchar](8000) NULL,
					@Payee [varchar](255) NULL,
					@ChequeNo [varchar] (50) NULL,
					@BankAccNo [varchar] (50) NULL,
					@ChqBookNo int NULL,
					@DueDate [datetime] NULL,
					@Handed_Over_To [varchar] (255) NULL
				)
AS
BEGIN
	 
	SET NOCOUNT ON;

	DECLARE @VchrType VARCHAR(5)
	SET @VchrType=SUBSTRING(@VchrNo,1,CHARINDEX('-',@VchrNo,1)-1)
	---
	DECLARE @NewVchrSequence INT; -- Holds the sequential part (e.g., 1, 2, 3)
    DECLARE @VchrNoPrefix VARCHAR(255); -- Holds the fixed prefix (e.g., 'PAY-251126-')
    DECLARE @FinalVchrNo VARCHAR(255); -- Holds the final, complete VchrNo string

    -- 1. CALCULATE THE PREFIX
    SET @VchrNoPrefix = @VchrType + '-' + FORMAT(@DT, 'yyMMdd') + '-';

    -- START THE DATABASE TRANSACTION AND APPLY LOCKING
    BEGIN TRANSACTION;
    
    SELECT 
        @NewVchrSequence = ISNULL(MAX(CAST(RIGHT(VchrNo,CHARINDEX('-',REVERSE(VchrNo))-1) AS INT)), 0) + 1
    FROM 
        dbo.Vouchers WITH (UPDLOCK, HOLDLOCK) -- LOCKING IS APPLIED HERE
    WHERE 
        VDate = @DT 
        AND LEFT(VchrNo, 8 + LEN(@VchrType)) = @VchrNoPrefix;

    -- 3. ASSEMBLE THE FINAL UNIQUE VOUCHER NUMBER
    SET @FinalVchrNo = @VchrNoPrefix + FORMAT(@NewVchrSequence, '000');


	IF @ChequeNo IS NOT NULL AND @BankAccNo IS NOT NULL
	BEGIN
		DECLARE @IsAvailable BIT
		SELECT @IsAvailable=Issued FROM dbo.VChqList WITH (UPDLOCK, HOLDLOCK)		
			WHERE AccNo=@BankAccNo
			AND ChqNo=@ChequeNo
		IF @IsAvailable IS NULL OR @IsAvailable=1
		BEGIN
			ROLLBACK TRANSACTION;
			THROW 50001, N'The Cheque Number is already used.',1;
		END
	END

	INSERT INTO [dbo].[VoucherInfo]
           (
			   [VchrNo]
			   ,[UserName]
			   ,[MachineName]
			   ,[DT]           
			   ,[Notes]           
			   ,[Payee]			   
			   ,[DueDate]
			   ,[Handed_Over_To]
		   )
		OUTPUT INSERTED.VchrNo  --Return the Final VchrNo
     VALUES
		(
			@FinalVchrNo
			,@UserName
			,@MachineName
			,@DT           
			,@Notes           
			,@Payee			
			,@DueDate
			,@Handed_Over_To
		)

	IF @ChequeNo IS NOT NULL AND @BankAccNo IS NOT NULL
		BEGIN
			UPDATE ChqList SET ChqList.Issued=1
				WHERE ChqList.ChqBookNo=@ChqBookNo AND ChqList.ChqNo=@ChequeNo
		END

	COMMIT TRANSACTION;

END


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[InsertVoucherLineItems_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[InsertVoucherLineItems_SP];
END
GO

CREATE   PROCEDURE [dbo].[InsertVoucherLineItems_SP]
    @Vouchers_TT dbo.Vouchers_TT READONLY
AS
BEGIN
    
    SET NOCOUNT ON;

    INSERT INTO [dbo].[Vouchers] 
    (
        SNo,
		VDate,
		VchrNo,
        AccNo,
        Description,
		Debit,
		Credit,
		Balance,
		DpstSlip,
		CSNo
    )
    SELECT
		VL.SNo,
		VL.VDate,
        VL.VchrNo,
        VL.AccNo,
        VL.Description,
		VL.Debit,
        VL.Credit,
		--VL.Balance,		
		0,
		VL.DpstSlip,		
        VL.CSNo
        
    FROM @Vouchers_TT AS VL; 
	-- 1. Calculate and store the running balance ONE TIME
	SELECT
		V.SNo,
		V.AccNo,
		NewBalance = Accounts.OpenBal+SUM(V.Debit - V.Credit) OVER (PARTITION BY V.AccNo ORDER BY V.SNo)
	INTO #RecalcResults -- Stores the results in a temporary table
	FROM Vouchers V
	INNER JOIN Accounts ON V.AccNo=Accounts.AccNo
	WHERE V.AccNo IN (SELECT DISTINCT AccNo FROM @Vouchers_TT);
	--Update Running Balance in Vouchers
    UPDATE V
		SET V.Balance = R.NewBalance
		FROM Vouchers V
		JOIN #RecalcResults R ON V.SNo = R.SNo AND V.AccNo = R.AccNo;
	--Update Final Balance in Accounts Table
	UPDATE A
		SET A.Balance = X.FinalBalance
		FROM Accounts A
		INNER JOIN (
			SELECT r.AccNo, r.NewBalance AS FinalBalance
			FROM #RecalcResults r
			INNER JOIN (
				SELECT AccNo, MAX(SNo) AS LastSNo
				FROM #RecalcResults
				GROUP BY AccNo
			) L ON r.AccNo = L.AccNo AND r.SNo = L.LastSNo
		) X ON A.AccNo = X.AccNo;

	DROP TABLE #RecalcResults

END


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS OFF;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[MakerBilling_RcvingWise_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[MakerBilling_RcvingWise_SP];
END
GO

CREATE PROCEDURE [dbo].[MakerBilling_RcvingWise_SP] (@MakerID AS INT=0,@DTFrom AS DATETIME,@DTTo AS DATETIME,@CatID AS VARCHAR(255),@ProcessID AS INT=0,@MasterPONo VARCHAR(50)='')
AS
BEGIN

CREATE TABLE #myTBL (CatID VARCHAR(50))
CREATE TABLE #Processes(ProcessID INT)
IF @ProcessID=0 
	INSERT INTO #Processes SELECT ProcessID FROM Processes
ELSE
	INSERT INTO #Processes SELECT ProcessID FROM Processes WHERE ProcessID=@ProcessID
	
DECLARE @myQuery AS VARCHAR(1000)

IF @CatID='0' 
	INSERT INTO #myTBL SELECT CatID FROM ItemCatagories
ELSE
	BEGIN	
		SET @myQuery='INSERT INTO #myTBL(CatID) SELECT CatID FROM ItemCatagories WHERE CatID IN(' + @CatID +')'
		EXEC(@myQuery)
	END
	
SELECT Makers.VendID,Makers.VendID1,Makers.VenderName,VendIssued.EntryID,VendIssued.RecieptID,Vendissued.DT,VendIssued.ItemID,TabIssd.IssQty,TabIssd.Rate,TempAdvance,TabRcvd.RcvdQty ,RcvdQtyUnAuth,RcvdWeight,
RcvdWeightUnAuth,RcvdValue,RcvdValueUnAuth
--,(TabRcvd.RcvdQty+RcvdQtyUnAuth)/CAST(TabIssd.IssQty AS FLOAT)*100 AS PercentRcvd
,100 AS PercentRcvd
,TabRcvd.RepairAmt,ItemName,ItemNameUrdu,Wastage,Accounts.Balance As AccBalance,TabRcvd.OrderNo,
Issue_RefID,TabRcvd.EntryID AS RcvDetailEntryID,TabRcvd.RecieptID AS RcvRecieptID,TabRcvd.LotNo,
BillingProcRcvdQty-ISNULL(WastageQtyFromRepair,0) AS BillingProcRcvdQty,BillingProcAmount-(ISNULL(WastageQtyFromRepair,0)*TabIssd.Rate) AS BillingProcAmount
--BillingProcRcvdQty-ISNULL(WastageQtyFromRepair,0) AS BillingProcRcvdQty,BillingProcAmount-ISNULL(WastageQtyFromRepair*Rate,0) AS BillingProcAmount
--TabRcvd.RcvdQty-ISNULL(ReWorkQty,0)-ISNULL(WastageQtyFromRepair,0) AS BillingProcRcvdQty,((TabRcvd.RcvdQty-ISNULL(ReWorkQty,0)-ISNULL(WastageQtyFromRepair,0))*Rate) AS BillingProcAmount
,OnlyDT,TabRcvd.ProcessID,Processes.Description,
CatID,TabRcvd.MasterPONo,MasterRecieptID,MainIssuanceEntryID,LotClearDT
,WastageQtyFromRepair,WastageQtyFromRepair*VendAssItems.Rate AS WastageValueFromRepair,RcvDT,ItemGroups.Description AS ItemGroupName,Audited,BillQtyDirect,BillQtyInDirect,FCustomerOrders.CustCode
,VendIssued.BillingDays,TRunningLots.Lot_Current_Process
,FCustomerOrders.InternalRefNo
FROM VendIssued
LEFT OUTER JOIN VendIssdDetail TabIssd ON VendIssued.EntryID=TabIssd.RefID
LEFT OUTER JOIN VendAssItems ON VendIssued.VendID=VendAssItems.VendID AND VendIssued.ProcessID=VendAssItems.ProcessID AND TabIssd.ItemCode=VendAssItems.ItemID
LEFT OUTER JOIN (SELECT VVendRcvdDetail.EntryID,Issuance_RefID,Issue_RefID,OrderNo,0 AS Not_Available_For_Billing
					--,CASE WHEN ReqAuth=0 THEN RcvdQty ELSE 0 END +ISNULL(SplitQtyForRcvdQty,0) AS RcvdQty
					,CASE WHEN ReqAuth=0 THEN RcvdQty ELSE 0 END AS RcvdQty
					,CASE WHEN ReqAuth=1 THEN RcvdQty ELSE 0 END AS RcvdQtyUnAuth,
					CASE WHEN ReqAuth=0 THEN RcvdWeight ELSE 0 END AS RcvdWeight,CASE WHEN ReqAuth=1 THEN RcvdWeight ELSE 0 END AS RcvdWeightUnAuth,
					CASE WHEN ReqAuth=0 THEN RcvdValue_For_Billing ELSE 0 END AS RcvdValue,CASE WHEN ReqAuth=1 THEN RcvdValue_For_Billing ELSE 0 END AS RcvdValueUnAuth,RepairAmt AS RepairAmt
,Wastage,VVendRcvdDetail.LotNo,VVendRcvdDetail.RecieptID,ISNULL(BillQtyDirect,0)+ISNULL(BillQtyInDirect,0) AS BillingProcRcvdQty,(ISNULL(BillQtyDirect,0)+ISNULL(BillQtyInDirect,0))*IssRate AS BillingProcAmount,OnlyDT,VVendRcvdDetail.ProcessID,MasterPONo,TMasters.RecieptID AS MasterRecieptID,MainIssuanceEntryID,LotClearDT,DT AS RcvDT,ReWorkQty,Audited,BillQtyDirect,BillQtyInDirect
FROM VVendRcvdDetail 
LEFT OUTER JOIN(SELECT VRD_From_RefID,SUM(SplitQty) AS SplitQtyForRcvdQty FROM LotTransferDetails GROUP BY VRD_From_RefID) TSplitQtyForRcvdQty ON VVendRcvdDetail.EntryID=TSplitQtyForRcvdQty.VRD_From_RefID
/*INNER JOIN
(SELECT LotNo,ProcessID,MAX(DT) AS LotClearDT,SUM(BillingProcRcvdQty) AS BillingProcRcvdQty1 FROM(
SELECT LotNo,VendRcvdDetail.ProcessID,DT,RcvdQty-ISNULL(ReWorkQty,0)-ISNULL(Wastage,0) AS BillingProcRcvdQty FROM VendRcvdDetail INNER JOIN VendReceived ON VendReceived.EntryID=VendRcvdDetail.RefID WHERE (VendID=@MakerID OR @MakerID=0) AND ReqAuth=0
UNION ALL
SELECT VRD_From.LotNo,VRD_Lots.ProcessID,VendReceived.DT,VRD_Lots.RcvdQty-ISNULL(VRD_Lots.ReWorkQty,0) AS BillingProcRcvdQty FROM LotTransferDetails
	INNER JOIN VendRcvdDetail VRD_From ON LotTransferDetails.VRD_From_RefID=VRD_From.EntryID
	INNER JOIN VendRcvdDetail VRD_To ON LotTransferDetails.VRD_To_RefID=VRD_To.EntryID
	INNER JOIN VendRcvdDetail VRD_Lots ON VRD_To.LotNo=VRD_Lots.LotNo
	INNER JOIN VendReceived ON VRD_Lots.RefID=VendReceived.EntryID
	WHERE VRD_Lots.ReqAuth=0 AND Type=1 --AND VRD_To.ProcessID <> VendRcvdDetail.ProcessID
	) TBilling GROUP BY LotNo,ProcessID
) TBillProc ON VVendRcvdDetail.BillingProcessID=TBillProc.ProcessID AND VVendRcvdDetail.LotNo=TBillProc.LotNo*/

INNER JOIN(SELECT LotNo,VendRcvdDetail.ProcessID,MAX(DT) AS LotClearDT,SUM(RcvdQty-ISNULL(ReWorkQty,0)-ISNULL(Wastage,0)) AS BillQtyDirect FROM VendRcvdDetail INNER JOIN VendReceived  ON VendReceived.EntryID=VendRcvdDetail.RefID GROUP BY LotNo,VendRcvdDetail.ProcessID) TBillQtyDirect ON VVendRcvdDetail.LotNo=TBillQtyDirect.LotNo AND VVendRcvdDetail.BillingProcessID=TBillQtyDirect.ProcessID
LEFT OUTER JOIN(SELECT VRD_From.LotNo,VRD_Lots.ProcessID,SNo,SUM(VRD_Lots.RcvdQty-ISNULL(VRD_Lots.ReWorkQty,0)) AS BillQtyInDirect FROM LotTransferDetails
	INNER JOIN VendRcvdDetail VRD_From ON LotTransferDetails.VRD_From_RefID=VRD_From.EntryID
	INNER JOIN VendRcvdDetail VRD_To ON LotTransferDetails.VRD_To_RefID=VRD_To.EntryID
	INNER JOIN VendRcvdDetail VRD_Lots ON VRD_To.LotNo=VRD_Lots.LotNo
	INNER JOIN VendReceived ON VRD_Lots.RefID=VendReceived.EntryID
	INNER JOIN ItemProcesses ON VRD_From.ProcessID=ItemProcesses.ProcessID AND ItemProcesses.ItemID=VRD_From.ItemCode 
	WHERE VRD_Lots.ReqAuth=0 AND Type=1 GROUP BY VRD_From.LotNo,VRD_Lots.ProcessID,SNo) TBillQtyInDirect ON VVendRcvdDetail.LotNo=TBillQtyInDirect.LotNo AND VVendRcvdDetail.BillingProcessID=TBillQtyInDirect.ProcessID AND (TBillQtyInDirect.SNo>VVendRcvdDetail.ItemSNo)

LEFT OUTER JOIN (SELECT LotNo,MIN(EntryID) AS MinEntryID FROM VendRcvdDetail GROUP BY LotNo) TMinLot ON VVendRcvdDetail.LotNo=TMinLot.LotNo
LEFT OUTER JOIN (SELECT VendIssued.EntryID AS MainIssuanceEntryID,VendRcvdDetail.EntryID,MasterPONo,VendIssued.RecieptID FROM VendIssued INNER JOIN VendIssdDetail ON VendIssdDetail.RefID=VendIssued.EntryID INNER JOIN VendRcvdDetail ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID) TMasters
	ON TMinLot.MinEntryID=TMasters.EntryID

WHERE (VendID=@MakerID OR @MakerID=0) AND OnlyDT BETWEEN @DTFrom AND @DTTo  --AND ReWorkLot=0
			) TabRcvd ON VendIssued.EntryID=TabRcvd.Issuance_RefID AND TabIssd.EntryID=TabRcvd.Issue_RefID
LEFT OUTER JOIN (SELECT Issuance_RefID,SUM(Amount) AS TempAdvance FROM MakerTemporaryLoans WHERE ISNULL(GroupID,0)=0 OR GroupID IN(SELECT EntryID FROM MakerTemporaryLoansGroup WHERE Active=1) 
																	GROUP BY Issuance_RefID) TabTemp ON VendIssued.EntryID=TabTemp.Issuance_RefID

LEFT OUTER JOIN (SELECT VRD_RefID_ChargeTo,SUM(Qty) AS WastageQtyFromRepair FROM VMakerRepair_BIlling WHERE ShowInBilling=0 AND ItemSNo>BillingProcessSNo AND BillingProcessSNo IS NOT NULL GROUP BY VRD_RefID_ChargeTo) TMakerWastage ON TabRcvd.EntryID=TMakerWastage.VRD_RefID_ChargeTo
INNER JOIN Items ON VendIssued.ItemID=Items.ItemID
INNER JOIN ItemGroups ON Items.GroupID=ItemGroups.ID
INNER JOIN Makers ON VendIssued.VendID=Makers.VendID
INNER JOIN Accounts ON Makers.AccNo=Accounts.AccNo
INNER JOIN Processes ON TabRcvd.ProcessID=Processes.ProcessID
INNER JOIN #Processes ON #Processes.ProcessID=TabRcvd.ProcessID
INNER JOIN FCustomerOrders ON TabIssd.OrderNo=FCustomerOrders.OrderNo
LEFT JOIN (SELECT LotNo,MAX(Description) AS Lot_Current_Process FROM VRunningLots_Simple GROUP BY LotNo) TRunningLots ON TabRcvd.LotNo=TRunningLots.LotNo
WHERE (VendIssued.VendID=@MakerID OR @MakerID=0) AND Issue_RefID IS NOT NULL
AND TabRcvd.EntryID NOT IN(SELECT VRD_RefID FROM MakerPostedBillsDetail_Receivings)
AND Items.CatID IN(SELECT CatID FROM #myTBL)
AND TabRcvd.EntryID NOT IN(SELECT VRD_RefID FROM MakerBilling_Receivings_Hidden)
AND TabRcvd.Not_Available_For_Billing=0
AND ((TabRcvd.MasterPONo LIKE '%'+@MasterPONo+'%') OR @MasterPONo='')
AND OnlyDT>='12/1/2019'
--ORDER BY RcvDT
ORDER BY FCustomerOrders.InternalRefNo
--ORDER BY (RcvdQty+RcvdQtyUnAuth)/CAST(TabIssd.IssQty AS FLOAT) *100 DESC
END


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PrintItemLabelsCustomerWise_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[PrintItemLabelsCustomerWise_SP];
END
GO

CREATE   PROCEDURE [dbo].[PrintItemLabelsCustomerWise_SP] (@CustCode VARCHAR(50),@ItemID AS VARCHAR(255),@No AS INT,@PONo AS VARCHAR(255),@Qty AS INT) AS

BEGIN

	DECLARE @TempTB TABLE(AutoEntryID INT IDENTITY (1, 1),ItemNo INT,ItemID VARCHAR(255))

	DECLARE @ItemNo AS INT


	DECLARE @iCounter AS INT
	SET @iCounter=1
	WHILE @iCounter<=@No
	BEGIN
		INSERT INTO @TempTB(ItemNo,ItemID) VALUES(@iCounter,@ItemID)
		SET @iCounter=@iCounter+1
	END
			

END

SELECT TempTB.*,ItemName,@PONo AS PONo,@Qty AS Qty
	,FCustomerCatalog.ItemID AS Customer_ItemID	
	FROM @TempTB TempTB
	INNER JOIN Items ON Items.ItemID=TempTB.ItemID
	INNER JOIN FCustomerCatalog ON Items.ItemID=FCustomerCatalog.CompItemID AND FCustomerCatalog.CustCode=@CustCode

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ProductionOrderReport_WithSet_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[ProductionOrderReport_WithSet_SP];
END
GO


/****** Object:  StoredProcedure [dbo].[ProductionOrderReportPHF_SP]    Script Date: 03/10/2014 09:45:10 ******/
/*SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO*/

CREATE PROCEDURE [dbo].[ProductionOrderReport_WithSet_SP] (@OrderNo AS VARCHAR(255)) AS

BEGIN
	SET ANSI_NULLS ON
	
	SET ARITHABORT ON


	SELECT     dbo.VrptOrders_ForProduction.OrderNo, dbo.VrptOrders_ForProduction.DT, dbo.VrptOrders_ForProduction.TradeTerms, dbo.VrptOrders_ForProduction.CustCode, dbo.VrptOrders_ForProduction.Country, dbo.VrptOrders_ForProduction.PartialShipment, 
                      dbo.VrptOrders_ForProduction.PaymentTerms, dbo.VrptOrders_ForProduction.TransShipment, dbo.VrptOrders_ForProduction.Packaging, dbo.VrptOrders_ForProduction.DeliveryDT, dbo.VrptOrders_ForProduction.BatchNo, 
                      dbo.VrptOrders_ForProduction.CompanyRefID, dbo.VrptOrders_ForProduction.StampDT, dbo.VrptOrders_ForProduction.Quality, dbo.VrptOrders_ForProduction.InternalRefNo, dbo.VrptOrders_ForProduction.OrderRcvdVia, 0 AS ID, 
                      '' AS ItemCode, 0 AS Price, dbo.VrptOrders_ForProduction.Qty, 0 AS InvQty, dbo.VrptOrders_ForProduction.ItemName, dbo.VrptOrders_ForProduction.Description, 
                      dbo.VrptOrders_ForProduction.ItemID, dbo.VrptOrders_ForProduction.CompItemID, dbo.VrptOrders_ForProduction.Curr, dbo.VrptOrders_ForProduction.Unit, dbo.VrptOrders_ForProduction.Name, dbo.VrptOrders_ForProduction.Address, 
                      0 AS CustomPrice, dbo.VrptOrders_ForProduction.CustomDescription, dbo.VrptOrders_ForProduction.ItemSize, dbo.VrptOrders_ForProduction.SizeUnit, dbo.VrptOrders_ForProduction.SteelType, dbo.VrptOrders_ForProduction.Gage, 
                      dbo.VrptOrders_ForProduction.ItemColor, dbo.VrptOrders_ForProduction.BarcodeNo, dbo.VrptOrders_ForProduction.PackingMode, dbo.VrptOrders_ForProduction.MasterCartonL, dbo.VrptOrders_ForProduction.MasterCartonW, 
                      dbo.VrptOrders_ForProduction.MasterCartonH, dbo.VrptOrders_ForProduction.SmallCartonL, dbo.VrptOrders_ForProduction.SmallCartonW, dbo.VrptOrders_ForProduction.SmallCartonH, dbo.VrptOrders_ForProduction.PolyBag, 
                      dbo.VrptOrders_ForProduction.SpecialInstructions, dbo.VrptOrders_ForProduction.StampInstructions, dbo.VrptOrders_ForProduction.PackingInstructions, dbo.VrptOrders_ForProduction.UnitItems, dbo.VrptOrders_ForProduction.CustomerColor, 
                      dbo.VrptOrders_ForProduction.ShippedQty, dbo.VrptOrders_ForProduction.ItemNameUrdu, dbo.VrptOrders_ForProduction.InHand, 0 AS SortNo, dbo.VrptOrders_ForProduction.MaxLotSize, dbo.VrptOrders_ForProduction.FinQuality
                      ,dbo.VrptOrders_ForProduction.Type, TabPO.RecieptID, TabPO.IssQty, TabAdj.AdjQty
                      ,TPTC.PTC
                      ,TProcesses.Processes
                      ,TPOQty.POQty
                      ,TPONo.PONos
                      ,TabPacked.PackQty
                      ,TCartons.Cartons
                      ,TabQC.RcvdQtyQC,TabQC.WastageQC,TabQC.ReWorkQtyQC,TabInsp.RcvdQtyInsp,TabInsp.WastageInsp,TabInsp.ReWorkQtyInsp,ReWorkQtyQCBalance,TotalQtyRcvd,TotalTransferred,ReWorkQtyInspBalance
                      ,TIMR.InitialMakerRcvings
                      ,'' AS DeliveryDTItem,'' AS DeliveryStatus,'' AS StampsItem,'' AS QualityItem,'' AS Remarks
					  ,TPOMN.POMakerName
					  ,VenderName,ReturnDT,'' AS GroupDescription,0 AS GroupID,CompanyName,0 AS FinishedWeight,0 AS SmallBoxPcs,0 AS MasterCartonSmallBoxes
					  ,0 AS MainItem,ItemsSets.Set_ItemID
FROM         dbo.VrptOrders_ForProduction WITH (NOLOCK) 
			LEFT OUTER JOIN Companies WITH (NOLOCK)ON VrptOrders_ForProduction.CompanyRefID=Companies.EntryID 
			LEFT OUTER JOIN ItemsSets WITH (NOLOCK)ON VrptOrders_ForProduction.ItemID=ItemsSets.ItemID
			AND VrptOrders_ForProduction.CompItemID=ItemsSets.Set_ItemID
			LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, RecieptID, IssQty,VenderName,ReturnDT
                            FROM          dbo.VVendIssdDetail WITH (NOLOCK) INNER JOIN
                            (SELECT     MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo AND VendID<>653 GROUP BY OrderNo, ItemCode) TI1 ON dbo.VVendIssdDetail.EntryID=TI1.MinEntryID) TabPO ON dbo.VrptOrders_ForProduction.OrderNo = TabPO.OrderNo AND 
                      dbo.VrptOrders_ForProduction.CompItemID = TabPO.ItemCode 
					  
					  LEFT OUTER JOIN
                      (SELECT     OrderNo, ItemID, SUM(Qty) AS AdjQty
                      FROM          StockOrderAdjustments WITH (NOLOCK) WHERE OrderNo=@OrderNo
                      GROUP BY OrderNo, ItemID) TabAdj ON dbo.VrptOrders_ForProduction.OrderNo = TabAdj.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TabAdj.ItemID LEFT OUTER JOIN
                      (SELECT     OrderNo, ItemCode, SUM(Qty) AS PackQty
                      FROM          VDispatchListDetail  WITH (NOLOCK) WHERE OrderNo=@OrderNo
                      GROUP BY OrderNo, ItemCode) TabPacked ON dbo.VrptOrders_ForProduction.OrderNo = TabPacked.OrderNo AND 
                      dbo.VrptOrders_ForProduction.CompItemID = TabPacked.ItemCode 
					  
					  LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(RcvdQty) AS RcvdQtyQC, SUM(Wastage + LostQty) AS WastageQC, SUM(VendRcvdDetailReWorkDetail.Qty) AS ReWorkQtyQC, 
                                                   SUM(VendRcvdDetailReWorkDetail.Qty - VendRcvdDetailReWorkDetail.IssQty) AS ReWorkQtyQCBalance
                            FROM          VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN
                                                   VendRcvdDetailReWorkDetail  WITH (NOLOCK) ON VendRcvdDetail.EntryID = VendRcvdDetailReWorkDetail.VRD_RefID
                            WHERE      ProcessID = 69 AND ReqAuth = 0 AND ReWorkLot=0 AND OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabQC ON dbo.VrptOrders_ForProduction.OrderNo = TabQC.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TabQC.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(RcvdQty) AS RcvdQtyInsp, SUM(Wastage + LostQty) AS WastageInsp, SUM(VendRcvdDetailReWorkDetail.Qty) AS ReWorkQtyInsp,
			SUM(VendRcvdDetailReWorkDetail.Qty - VendRcvdDetailReWorkDetail.IssQty) AS ReWorkQtyInspBalance
                            FROM          VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN
                                                   VendRcvdDetailReWorkDetail  WITH (NOLOCK) ON VendRcvdDetail.EntryID = VendRcvdDetailReWorkDetail.VRD_RefID
                            WHERE      ProcessID = 65 AND ReqAuth = 0 AND ReWorkLot=0 AND OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabInsp ON dbo.VrptOrders_ForProduction.OrderNo = TabInsp.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TabInsp.ItemCode LEFT OUTER JOIN
		(SELECT VendRcvdDetail.OrderNo,VendRcvdDetail.ItemCode,SUM(VendRcvdDetail.RcvdQty) AS TotalQtyRcvd FROM VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN VendIssdDetail  WITH (NOLOCK) ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID 
			WHERE ISNULL(VendIssdDetail.LotNo,'0')='0' AND VendRcvdDetail.OrderNo=@OrderNo--VendRcvdDetail.EntryID IN(SELECT MIN(EntryID) FROM VendRcvdDetail WHERE ReWorkLot=0)  
			GROUP BY VendRcvdDetail.OrderNo,VendRcvdDetail.ItemCode) TTotalRcvd ON dbo.VrptOrders_ForProduction.OrderNo = TTotalRcvd.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TTotalRcvd.ItemCode LEFT OUTER JOIN
		(SELECT OrderNo,ItemCode,SUM(Qty) AS TotalTransferred FROM TransferredToReadyFinishLotsDetail  WITH (NOLOCK) INNER JOIN VendRcvdDetail  WITH (NOLOCK) ON TransferredToReadyFinishLotsDetail.VRD_RefID=VendRcvdDetail.EntryID WHERE OrderNo=@OrderNo
			GROUP BY OrderNo,ItemCode) TTTRF ON dbo.VrptOrders_ForProduction.OrderNo = TTTRF.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TTTRF.ItemCode 
			
			
		LEFT OUTER JOIN (SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT ','+LotNo + ' ('+CAST(RcvdQty AS VARCHAR(5)) +')' FROM VendRcvdDetail WITH (NOLOCK) 
	INNER JOIN (SELECT MAX(EntryID) AS MAXEntryID FROM VendRcvdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo GROUP BY LotNo) T1 ON VendRcvdDetail.EntryID=T1.MaxEntryID
	WHERE OrderNo=@OrderNo AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=FOrderItems.CompItemCode  ORDER BY LotNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS PTC
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPTC ON VrptOrders_ForProduction.OrderNo=TPTC.OrderNo AND VrptOrders_ForProduction.ItemID=TPTC.ItemCode
		

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+Description + + ' (' + CONVERT(VARCHAR(50),DT,5) + ') ('  + CAST(RcvdQty-ISNULL(Wastage,0)-ISNULL(ReWorkQty,0) AS VARCHAR(10)) + ')' FROM VendRcvdDetail WITH (NOLOCK) INNER JOIN VendReceived WITH (NOLOCK) ON VendReceived.EntryID=VendRcvdDetail.RefID 
	INNER JOIN Processes ON VendReceived.ProcessID=Processes.ProcessID
	INNER JOIN (SELECT MAX(EntryID) AS MAXEntryID FROM VendRcvdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=FOrderItems.CompItemCode GROUP BY LotNo) T1 ON VendRcvdDetail.EntryID=T1.MAXEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS Processes
	FROM FOrderItems WHERE OrderNo=@OrderNo) TProcesses ON VrptOrders_ForProduction.OrderNo=TProcesses.OrderNo AND VrptOrders_ForProduction.ItemID=TProcesses.ItemCode

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,	
(STUFF((SELECT CHAR(13) + CHAR(10)+CAST(IssQty AS VARCHAR(5)) FROM VVendIssdDetail_Simple WITH (NOLOCK) INNER JOIN 
	(SELECT MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE OrderNo=@OrderNo AND ProcessID IN(172,212) 
	AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VVendIssdDetail_Simple.EntryID=T1.MINEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS POQty
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPOQty ON VrptOrders_ForProduction.OrderNo=TPOQty.OrderNo AND VrptOrders_ForProduction.ItemID=TPOQty.ItemCode
	
LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,	
(STUFF((SELECT CHAR(13) + CHAR(10)+RecieptID FROM VVendIssdDetail_Simple WITH (NOLOCK) INNER JOIN 
	(SELECT MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE OrderNo=@OrderNo AND ProcessID IN(172,212) 
	AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VVendIssdDetail_Simple.EntryID=T1.MINEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS PONos
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPONo ON VrptOrders_ForProduction.OrderNo=TPONo.OrderNo AND VrptOrders_ForProduction.ItemID=TPONo.ItemCode
	
LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,FOrderItems.CompItemCode,
(STUFF((SELECT CHAR(13)+CHAR(10)+CAST(CartonNo AS VARCHAR(10)) + '(' + CAST(SUM(Qty) AS VARCHAR(10)) + ')' FROM VDispatchListDetail_Adv WITH (NOLOCK) WHERE OrderNo=@OrderNo
	AND VDispatchListDetail_Adv.OrderNo=FOrderItems.OrderNo AND VDispatchListDetail_Adv.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,CartonNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS Cartons
	FROM FOrderItems WHERE OrderNo=@OrderNo) TCartons ON VrptOrders_ForProduction.OrderNo=TCartons.OrderNo AND VrptOrders_ForProduction.CompItemID=TCartons.CompItemCode

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+VendRcvdDetail.LotNo+ ' (' + CAST(SUM(VendRcvdDetail.RcvdQty) AS VARCHAR(10)) + ')' FROM VendRcvdDetail WITH (NOLOCK) 
	INNER JOIN VendIssdDetail ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID WHERE VendIssdDetail.LotNo='0' AND VendRcvdDetail.ReWorkLot=0 AND VendRcvdDetail.EntryID NOT IN(SELECT EntryID FROM LotTransferDetails)
		AND VendRcvdDetail.OrderNo=@OrderNo	
		AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=ForderItems.CompItemCode
		GROUP BY VendRcvdDetail.LotNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS InitialMakerRcvings
	FROM FOrderItems WHERE OrderNo=@OrderNo) TIMR ON VrptOrders_ForProduction.OrderNo=TIMR.OrderNo AND VrptOrders_ForProduction.ItemID=TIMR.ItemCode
	
LEFT OUTER JOIN (SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+CHAR(10)+VenderName FROM VendIssdDetail WITH (NOLOCK) INNER JOIN VendIssued WITH (NOLOCK) ON VendIssued.EntryID=VendIssdDetail.RefID INNER JOIN Makers WITH (NOLOCK) ON Makers.VendID=VendIssued.VendID INNER JOIN
	(SELECT MIN(EntryID) MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE ProcessID IN(172,212) AND VVendIssdDetail_Simple.OrderNo=@OrderNo AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=ForderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VendIssdDetail.EntryID=T1.MinEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS POMakerName
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPOMN ON VrptOrders_ForProduction.OrderNo=TPOMN.OrderNo AND VrptOrders_ForProduction.ItemID=TPOMN.ItemCode
	WHERE VrptOrders_ForProduction.OrderNo=@OrderNo

END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ProductionOrderReportPHF_SP_Old]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[ProductionOrderReportPHF_SP_Old];
END
GO

CREATE   PROCEDURE [dbo].[ProductionOrderReportPHF_SP_Old] (@OrderNo AS VARCHAR(255)) AS

BEGIN
	SET ANSI_NULLS ON
	
	SET ARITHABORT ON
	

	SELECT     dbo.VrptOrders.OrderNo, dbo.VrptOrders.DT, dbo.VrptOrders.TradeTerms, dbo.VrptOrders.CustCode, dbo.VrptOrders.Country, dbo.VrptOrders.PartialShipment, 
                      dbo.VrptOrders.PaymentTerms, dbo.VrptOrders.TransShipment, dbo.VrptOrders.Packaging, dbo.VrptOrders.DeliveryDT, dbo.VrptOrders.BatchNo, 
                      dbo.VrptOrders.CompanyRefID, dbo.VrptOrders.StampDT, dbo.VrptOrders.Quality, dbo.VrptOrders.InternalRefNo, dbo.VrptOrders.OrderRcvdVia, dbo.VrptOrders.ID, 
                      dbo.VrptOrders.ItemCode, dbo.VrptOrders.Price, dbo.VrptOrders.Qty, dbo.VrptOrders.InvQty, dbo.VrptOrders.ItemName, dbo.VrptOrders.Description, 
                      dbo.VrptOrders.ItemID, dbo.VrptOrders.CompItemID, dbo.VrptOrders.Curr, dbo.VrptOrders.Unit, dbo.VrptOrders.Name, dbo.VrptOrders.Address, 
                      dbo.VrptOrders.CustomPrice, dbo.VrptOrders.CustomDescription, dbo.VrptOrders.ItemSize, dbo.VrptOrders.SizeUnit, dbo.VrptOrders.SteelType, dbo.VrptOrders.Gage, 
                      dbo.VrptOrders.ItemColor, dbo.VrptOrders.BarcodeNo, dbo.VrptOrders.PackingMode, dbo.VrptOrders.MasterCartonL, dbo.VrptOrders.MasterCartonW, 
                      dbo.VrptOrders.MasterCartonH, dbo.VrptOrders.SmallCartonL, dbo.VrptOrders.SmallCartonW, dbo.VrptOrders.SmallCartonH, dbo.VrptOrders.PolyBag, 
                      dbo.VrptOrders.SpecialInstructions, dbo.VrptOrders.StampInstructions, dbo.VrptOrders.PackingInstructions, dbo.VrptOrders.UnitItems, dbo.VrptOrders.CustomerColor, 
                      dbo.VrptOrders.ShippedQty, dbo.VrptOrders.ItemNameUrdu, dbo.VrptOrders.InHand, dbo.VrptOrders.SortNo, dbo.VrptOrders.MaxLotSize, dbo.VrptOrders.FinQuality
                      ,dbo.VrptOrders.Type, TabPO.RecieptID, TabPO.IssQty, TabAdj.AdjQty
                      ,TPTC.PTC
                      ,TProcesses.Processes
                      ,TPOQty.POQty
                      ,TPONo.PONos
                      ,TabPacked.PackQty
                      ,TCartons.Cartons
                      ,TabQC.RcvdQtyQC,TabQC.WastageQC,TabQC.ReWorkQtyQC,TabInsp.RcvdQtyInsp,TabInsp.WastageInsp,TabInsp.ReWorkQtyInsp,ReWorkQtyQCBalance,TotalQtyRcvd,TotalTransferred,ReWorkQtyInspBalance
                      ,TIMR.InitialMakerRcvings
                      ,DeliveryDTItem,DeliveryStatus,StampsItem,QualityItem,Remarks
					  ,TPOMN.POMakerName
					  ,VenderName,ReturnDT,GroupDescription,GroupID,CompanyName,FinishedWeight,SmallBoxPcs,MasterCartonSmallBoxes
					  --,dbo.VrptOrders.EmpID,dbo.VrptOrders.EmpName
					  ,VItemSFStock.Qty AS SFQty
					  ,VrptOrders.Item_Edited
					  ,VrptOrders.OrderRevisionNo
					  ,VrptOrders.ShippingMode
					  ,VrptOrders.Item_Finishing_Type_Text
FROM         dbo.VrptOrders WITH (NOLOCK) 
LEFT OUTER JOIN VItemSFStock ON dbo.VrptOrders.CompItemCode=VItemSFStock.ItemID
LEFT OUTER JOIN Companies WITH (NOLOCK)ON VrptOrders.CompanyRefID=Companies.EntryID LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, RecieptID, IssQty,VenderName,ReturnDT
                            FROM          dbo.VVendIssdDetail WITH (NOLOCK) INNER JOIN
                            (SELECT     MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo AND VendID<>653 GROUP BY OrderNo, ItemCode) TI1 ON dbo.VVendIssdDetail.EntryID=TI1.MinEntryID) TabPO ON dbo.VrptOrders.OrderNo = TabPO.OrderNo AND 
                      dbo.VrptOrders.CompItemID = TabPO.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemID, SUM(Qty) AS AdjQty
                            FROM          StockOrderAdjustments WITH (NOLOCK) WHERE OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemID) TabAdj ON dbo.VrptOrders.OrderNo = TabAdj.OrderNo AND dbo.VrptOrders.CompItemID = TabAdj.ItemID LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(Qty) AS PackQty
                            FROM          VDispatchListDetail  WITH (NOLOCK) WHERE OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabPacked ON dbo.VrptOrders.OrderNo = TabPacked.OrderNo AND 
                      dbo.VrptOrders.CompItemID = TabPacked.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(RcvdQty) AS RcvdQtyQC, SUM(Wastage + LostQty) AS WastageQC, SUM(VendRcvdDetailReWorkDetail.Qty) AS ReWorkQtyQC, 
                                                   SUM(VendRcvdDetailReWorkDetail.Qty - VendRcvdDetailReWorkDetail.IssQty) AS ReWorkQtyQCBalance
                            FROM          VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN
                                                   VendRcvdDetailReWorkDetail  WITH (NOLOCK) ON VendRcvdDetail.EntryID = VendRcvdDetailReWorkDetail.VRD_RefID
                            WHERE      ProcessID = 69 AND ReqAuth = 0 AND ReWorkLot=0 AND OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabQC ON dbo.VrptOrders.OrderNo = TabQC.OrderNo AND dbo.VrptOrders.CompItemID = TabQC.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(RcvdQty) AS RcvdQtyInsp, SUM(Wastage + LostQty) AS WastageInsp, SUM(VendRcvdDetailReWorkDetail.Qty) AS ReWorkQtyInsp,
			SUM(VendRcvdDetailReWorkDetail.Qty - VendRcvdDetailReWorkDetail.IssQty) AS ReWorkQtyInspBalance
                            FROM          VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN
                                                   VendRcvdDetailReWorkDetail  WITH (NOLOCK) ON VendRcvdDetail.EntryID = VendRcvdDetailReWorkDetail.VRD_RefID
                            WHERE      ProcessID = 65 AND ReqAuth = 0 AND ReWorkLot=0 AND OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabInsp ON dbo.VrptOrders.OrderNo = TabInsp.OrderNo AND dbo.VrptOrders.CompItemID = TabInsp.ItemCode LEFT OUTER JOIN
		(SELECT VendRcvdDetail.OrderNo,VendRcvdDetail.ItemCode,SUM(VendRcvdDetail.RcvdQty) AS TotalQtyRcvd FROM VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN VendIssdDetail  WITH (NOLOCK) ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID 
			WHERE ISNULL(VendIssdDetail.LotNo,'0')='0' AND VendRcvdDetail.OrderNo=@OrderNo--VendRcvdDetail.EntryID IN(SELECT MIN(EntryID) FROM VendRcvdDetail WHERE ReWorkLot=0)  
			GROUP BY VendRcvdDetail.OrderNo,VendRcvdDetail.ItemCode) TTotalRcvd ON dbo.VrptOrders.OrderNo = TTotalRcvd.OrderNo AND dbo.VrptOrders.CompItemID = TTotalRcvd.ItemCode LEFT OUTER JOIN
		(SELECT OrderNo,ItemCode,SUM(Qty) AS TotalTransferred FROM TransferredToReadyFinishLotsDetail  WITH (NOLOCK) INNER JOIN VendRcvdDetail  WITH (NOLOCK) ON TransferredToReadyFinishLotsDetail.VRD_RefID=VendRcvdDetail.EntryID WHERE OrderNo=@OrderNo
			GROUP BY OrderNo,ItemCode) TTTRF ON dbo.VrptOrders.OrderNo = TTTRF.OrderNo AND dbo.VrptOrders.CompItemID = TTTRF.ItemCode 
			
			
		LEFT OUTER JOIN (SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT ','+LotNo + ' ('+CAST(RcvdQty AS VARCHAR(5)) +')' FROM VendRcvdDetail WITH (NOLOCK) 
	INNER JOIN (SELECT MAX(EntryID) AS MAXEntryID FROM VendRcvdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo GROUP BY LotNo) T1 ON VendRcvdDetail.EntryID=T1.MaxEntryID
	WHERE OrderNo=@OrderNo AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=FOrderItems.CompItemCode  ORDER BY LotNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS PTC
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPTC ON VrptOrders.OrderNo=TPTC.OrderNo AND VrptOrders.ItemID=TPTC.ItemCode
		

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+Description + + ' (' + CONVERT(VARCHAR(50),DT,5) + ') ('  + CAST(RcvdQty-ISNULL(Wastage,0)-ISNULL(ReWorkQty,0) AS VARCHAR(10)) + ')' FROM VendRcvdDetail WITH (NOLOCK) INNER JOIN VendReceived WITH (NOLOCK) ON VendReceived.EntryID=VendRcvdDetail.RefID 
	INNER JOIN Processes ON VendReceived.ProcessID=Processes.ProcessID
	INNER JOIN (SELECT MAX(EntryID) AS MAXEntryID FROM VendRcvdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=FOrderItems.CompItemCode GROUP BY LotNo) T1 ON VendRcvdDetail.EntryID=T1.MAXEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS Processes
	FROM FOrderItems WHERE OrderNo=@OrderNo) TProcesses ON VrptOrders.OrderNo=TProcesses.OrderNo AND VrptOrders.ItemID=TProcesses.ItemCode

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,	
(STUFF((SELECT CHAR(13) + CHAR(10)+CAST(IssQty AS VARCHAR(5)) FROM VVendIssdDetail_Simple WITH (NOLOCK) INNER JOIN 
	(SELECT MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE OrderNo=@OrderNo AND MasterPONo<>''
	AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VVendIssdDetail_Simple.EntryID=T1.MINEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS POQty
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPOQty ON VrptOrders.OrderNo=TPOQty.OrderNo AND VrptOrders.ItemID=TPOQty.ItemCode
	
	
	
LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,	
(STUFF((SELECT CHAR(13) + CHAR(10)+RecieptID FROM VVendIssdDetail_Simple WITH (NOLOCK) INNER JOIN 
	(SELECT MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE OrderNo=@OrderNo AND ProcessID IN(172,212) 
	AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VVendIssdDetail_Simple.EntryID=T1.MINEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS PONos
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPONo ON VrptOrders.OrderNo=TPONo.OrderNo AND VrptOrders.ItemID=TPONo.ItemCode
	
LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+CHAR(10)+CAST(CartonNo AS VARCHAR(10)) + '(' + CAST(SUM(Qty) AS VARCHAR(10)) + ')' FROM VDispatchListDetail_Adv WITH (NOLOCK) WHERE OrderNo=@OrderNo
	AND VDispatchListDetail_Adv.OrderNo=FOrderItems.OrderNo AND VDispatchListDetail_Adv.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,CartonNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS Cartons
	FROM FOrderItems WHERE OrderNo=@OrderNo) TCartons ON VrptOrders.OrderNo=TCartons.OrderNo AND VrptOrders.ItemID=TCartons.ItemCode

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+VendRcvdDetail.LotNo+ ' (' + CAST(SUM(VendRcvdDetail.RcvdQty) AS VARCHAR(10)) + ')' FROM VendRcvdDetail WITH (NOLOCK) 
	INNER JOIN VendIssdDetail ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID WHERE VendIssdDetail.LotNo='0' AND VendRcvdDetail.ReWorkLot=0 AND VendRcvdDetail.EntryID NOT IN(SELECT EntryID FROM LotTransferDetails)
		AND VendRcvdDetail.OrderNo=@OrderNo	
		AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=ForderItems.CompItemCode
		GROUP BY VendRcvdDetail.LotNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS InitialMakerRcvings
	FROM FOrderItems WHERE OrderNo=@OrderNo) TIMR ON VrptOrders.OrderNo=TIMR.OrderNo AND VrptOrders.ItemID=TIMR.ItemCode
	
LEFT OUTER JOIN (SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+CHAR(10)+VenderName FROM VendIssdDetail WITH (NOLOCK) INNER JOIN VendIssued WITH (NOLOCK) ON VendIssued.EntryID=VendIssdDetail.RefID INNER JOIN Makers WITH (NOLOCK) ON Makers.VendID=VendIssued.VendID INNER JOIN
	(SELECT MIN(EntryID) MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE MasterPONo<>'' AND VVendIssdDetail_Simple.OrderNo=@OrderNo AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=ForderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VendIssdDetail.EntryID=T1.MinEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS POMakerName
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPOMN ON VrptOrders.OrderNo=TPOMN.OrderNo AND VrptOrders.ItemID=TPOMN.ItemCode
	WHERE VrptOrders.OrderNo=@OrderNo

END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PTC_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[PTC_SP];
END
GO

CREATE PROCEDURE [dbo].[PTC_SP] (@LotNo AS VARCHAR(50)) AS

BEGIN
--SET ANSI_WARNINGS OFF

DECLARE @RepairGeneratedFrom AS VARCHAR(4000),@RepairType AS VARCHAR(255)


--IF @Repair=0
	BEGIN
	SET @RepairGeneratedFrom=''
	SET @RepairType=''
SELECT     TFOrderItems.OrderNo, TFOrderItems.ItemCode, TFOrderItems.Qty, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.DT, 
                      dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.FCustomerOrders.Packaging, dbo.FCustomerOrders.DeliveryDT, 
                      dbo.ItemProcesses.SNO, dbo.ItemProcesses.ProcessID, dbo.Processes.Description, Tab1.RcvdQty, Tab2.LotNo, Tab1.VendID, dbo.Makers.VenderName, Tab1.RcvDT, 
                      TabIss.IssDT, TabMin.MinProcessID, Tab1.Wastage, Tab1.PassQty, dbo.Employees.empid, dbo.Employees.name, Tab1.ReWorkQty, TFOrderItems.CompItemCode, 
                      dbo.Processes.Code, dbo.Processes.ProcessNameUrdu, dbo.Items.ItemNameUrdu, EmpRcv.name AS RcvEmpName, Tab1.Rate, TMasterPO.MasterPONo
                      ,TSteelBatch.SteelBatchNo
					  --,TForgeBatch.ForgeBatchNo
					  ,TSteelBatch.ForgeBatchNo
					  , TSFRemarks.SFRemarks, TSplit.SplitQty, ISNULL(TSplit.Type, TSplitFrom.SplitTypeFrom) AS SplitType, 
                      TSplit.ToOrderNo, TSplit.ToLotNo, TSplitFrom.FromOrderNo, TSplitFrom.FromLotNo, TSplitFrom.SplitQtyFrom, TSplit.FromItemCode, TSplit.ToItemCode, 
                      TMasterPO.RecieptID, TIssFromSF.LotNo AS IssSFLotNo, TIssFromSF.IssSFRemarks, TIssFromSF.IssSFQty, TLastRcv.LotLastEntryID, 
                      dbo.VFOrderItemsWithShippedQtyItemCodeWise.ShippedQty, ISNULL(dbo.VLotWithPolishingMaker.VenderName, '') AS FromLotPolisher, 
                      ISNULL(dbo.VLotWithFirstProcessOfIssuance.VenderName, '') AS FromLotOriginator, Tab1.RcvTemperValue, TabIss.IssSpecialInstructions, 
                      dbo.FCustomerOrders.InternalRefNo, dbo.ProcessGroups.PGReference, Tab1.ReceivedEntryID, dbo.VendReceived_Employees_F(Tab1.ReceivedEntryID) 
                      AS Employees, TFOrderItems.Quality, dbo.GetAllRepairLots_F(Tab2.LotNo) AS RepairLots, TSplit.LotTransferRemarks, 
                      TSplitFrom.LotTransferRemarks AS LotTransferRemarksFrom, Tab1.UserName, Tab1.MachineName,ProcessGroupsProcesses.Hub_Name,VI_DT,VLotsWithMasterPONo.Batch_No AS VID_Batch_No,VLotsWithMasterPONo.MasterPODT
					  ,TIssFromSF.StoreName,TIssFromSF.Location
					  ,Lots_List.Batch_No AS Lots_List_Batch_No,Lots_List.Reference_LotNo,Lots_List.Mill_Certificate_No,Mill_Certificate_No_From_RM
					  ,Items.HRC_From,Items.HRC_To
					  ,Tab_TTRF.TTRFQty
					  ,Iss_MakerName,Makers.VendID1 
FROM         (SELECT     OrderNo, ItemCode, CompItemCode, Quality, SUM(Qty) AS Qty
                       FROM          dbo.VrptOrders_ForProduction_Simple 
                       GROUP BY OrderNo, ItemCode, CompItemCode, Quality) AS TFOrderItems LEFT OUTER JOIN
                      dbo.VFOrderItemsWithShippedQtyItemCodeWise ON TFOrderItems.OrderNo = dbo.VFOrderItemsWithShippedQtyItemCodeWise.OrderNo AND 
                      TFOrderItems.CompItemCode = dbo.VFOrderItemsWithShippedQtyItemCodeWise.CompItemCode INNER JOIN
                      dbo.Items ON TFOrderItems.CompItemCode = dbo.Items.ItemID INNER JOIN
                      dbo.ItemProcesses ON dbo.Items.ItemID = dbo.ItemProcesses.ItemID INNER JOIN
                      dbo.FCustomerOrders ON TFOrderItems.OrderNo = dbo.FCustomerOrders.OrderNo INNER JOIN
                      dbo.Processes ON dbo.ItemProcesses.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
                      dbo.ItemProcessGroups ON dbo.Items.ItemID = dbo.ItemProcessGroups.ItemID LEFT OUTER JOIN
                      dbo.ProcessGroups ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroups.EntryID LEFT OUTER JOIN
					  ProcessGroupsProcesses ON dbo.ProcessGroups.EntryID=ProcessGroupsProcesses.Group_RefID AND ItemProcesses.ProcessID=ProcessGroupsProcesses.Process_RefID
					  LEFT OUTER JOIN
                          (SELECT DISTINCT OrderNo, ItemCode, LotNo
                             FROM dbo.VendRcvdDetail WHERE LotNo=@LotNo) AS Tab2 ON TFOrderItems.OrderNo = Tab2.OrderNo AND TFOrderItems.CompItemCode = Tab2.ItemCode LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_2.ProcessID,VendRcvdDetail_2.LotNo,VendRcvdDetail_2.Rate,dbo.VendReceived.EmpID, 
                                                   dbo.VendReceived.TemperValue AS RcvTemperValue,VendRcvdDetail_2.RcvdQty,CASE WHEN VendRcvdDetail_2.ReqAuth = 1 THEN 0 ELSE VendRcvdDetail_2.RcvdQty END AS PassQty, 
                                                   dbo.VendReceived.VendID,dbo.VendReceived.DT AS RcvDT,VendRcvdDetail_2.Wastage,VendRcvdDetail_2.ReWorkQty, 
                                                   dbo.VendReceived.EntryID AS ReceivedEntryID, dbo.VendReceived.UserName, dbo.VendReceived.MachineName,VendRcvdDetail_2.EntryID AS VRD_EntryID,VendIssued.DT AS VI_DT,VendIssdDetail.Batch_No AS VID_Batch_No
                             FROM dbo.VendRcvdDetail AS VendRcvdDetail_2 
							 INNER JOIN dbo.VendReceived ON dbo.VendReceived.EntryID = VendRcvdDetail_2.RefID
							 LEFT JOIN VendIssdDetail ON VendRcvdDetail_2.Issue_RefID=VendIssdDetail.EntryID
							 LEFT JOIN VendIssued ON VendIssdDetail.RefID=VendIssued.EntryID
								WHERE VendRcvdDetail_2.LotNo=@LotNo) AS Tab1 ON Tab2.LotNo = Tab1.LotNo AND 
                      dbo.ItemProcesses.ProcessID = Tab1.ProcessID LEFT OUTER JOIN
                          (SELECT     dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID,
                                                   dbo.VendIssued.SpecialInstructions AS IssSpecialInstructions, MIN(dbo.VendIssued.DT) AS IssDT,MIN(VendIssued.VendID) AS Iss_VendID,MIN(Makers.VenderName) AS Iss_MakerName
                             FROM         dbo.VendIssued INNER JOIN
                                                   dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID LEFT OUTER JOIN
                                                   dbo.VendIssdDetail_MoreDetails ON dbo.VendIssdDetail.EntryID = dbo.VendIssdDetail_MoreDetails.VID_RefID
												   INNER JOIN Makers ON dbo.VendIssued.VendID=Makers.VendID
                             GROUP BY dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssued.SpecialInstructions) AS TabIss ON 
                      Tab2.LotNo = TabIss.LotNo AND dbo.ItemProcesses.ProcessID = TabIss.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_1.LotNo, MIN(VendReceived_1.ProcessID) AS MinProcessID
                             FROM         dbo.VendReceived AS VendReceived_1 INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON VendReceived_1.EntryID = VendRcvdDetail_1.RefID  WHERE VendRcvdDetail_1.LotNo=@LotNo
                             GROUP BY VendRcvdDetail_1.LotNo) AS TabMin ON Tab2.LotNo = TabMin.LotNo LEFT OUTER JOIN
                      dbo.Makers ON Tab1.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                      dbo.Employees ON TabIss.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                      dbo.Employees AS EmpRcv ON Tab1.EmpID = EmpRcv.empid LEFT OUTER JOIN
                          (SELECT     TNest1.LotNo, VendIssued_2.EntryID, VendIssued_2.MasterPONo, VendIssued_2.RecieptID
                             FROM         dbo.VendIssued AS VendIssued_2 INNER JOIN
                                                   dbo.VendIssdDetail AS VendIssdDetail_2 ON VendIssued_2.EntryID = VendIssdDetail_2.RefID INNER JOIN
                                                       (SELECT     LotNo, MIN(Issue_RefID) AS Issue_RefID
                                                          FROM         dbo.VendRcvdDetail AS VendRcvdDetail_7 WHERE LotNo=@LotNo
                                                          GROUP BY LotNo) AS TNest1 ON TNest1.Issue_RefID = VendIssdDetail_2.EntryID) AS TMasterPO ON 
                      Tab2.LotNo = TMasterPO.LotNo LEFT OUTER JOIN
                          (SELECT     dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS SteelBatchNo,MAX(VendRcvdDetailPO.RcvID) AS ForgeBatchNo
								,MAX(MaterialLocationwiseStatus.Mill_Certificate_No) AS Mill_Certificate_No_From_RM
                             FROM         dbo.RawMaterialIssuance INNER JOIN
                                                   dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                   dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
												   INNER JOIN RMID_MLS_Details ON RawMaterialIssuanceDetail.EntryID=RMID_MLS_Details.RMID_RefID
												   INNER JOIN MaterialLocationwiseStatus ON RMID_MLS_Details.MLS_RefID=MaterialLocationwiseStatus.EntryID
												   INNER JOIN VendRcvdDetailPO ON MaterialLocationwiseStatus.Rcvd_RefID=VendRcvdDetailPO.EntryID
                             WHERE     (dbo.RM.GroupID = 1)
                             GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TSteelBatch ON TMasterPO.EntryID = TSteelBatch.VI_RefID LEFT OUTER JOIN
                          /*(SELECT     RawMaterialIssuance_1.VI_RefID, MAX(RawMaterialIssuanceDetail_1.PORefNo) AS ForgeBatchNo
                             FROM         dbo.RawMaterialIssuance AS RawMaterialIssuance_1 INNER JOIN
                                                   dbo.RawMaterialIssuanceDetail AS RawMaterialIssuanceDetail_1 ON RawMaterialIssuance_1.IssNo = RawMaterialIssuanceDetail_1.IssNo INNER JOIN
                                                   dbo.RM AS RM_1 ON RawMaterialIssuanceDetail_1.RMID1 = RM_1.RMID1
                             WHERE     (RM_1.GroupID IN (25, 27, 28, 29, 30))
                             GROUP BY RawMaterialIssuance_1.VI_RefID) AS TForgeBatch ON TMasterPO.EntryID = TForgeBatch.VI_RefID LEFT OUTER JOIN*/
						
                          (SELECT     VendRcvdDetail_6.ProcessID, VendRcvdDetail_6.LotNo, '' AS SFRemarks
                             FROM         dbo.VendRcvdDetail AS VendRcvdDetail_6 INNER JOIN
                                                       (SELECT     VID_RefID
                                                          FROM         dbo.StockOrderOpening_Issuance
                                                          GROUP BY VID_RefID) AS TSOO_I ON TSOO_I.VID_RefID = VendRcvdDetail_6.Issue_RefID) AS TSFRemarks ON Tab2.LotNo = TSFRemarks.LotNo AND
                       Tab1.ProcessID = TSFRemarks.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_5.LotNo, VendRcvdDetail_5.ProcessID, dbo.LotTransferDetails.SplitQty, dbo.LotTransferDetails.Type, 
                                                   dbo.LotTransferDetails.ToOrderNo, VendRcvdDetail_1.LotNo AS ToLotNo, dbo.LotTransferDetails.FromItemCode, dbo.LotTransferDetails.ToItemCode, 
                                                   dbo.LotTransferDetails.LotTransferRemarks
                             FROM         dbo.LotTransferDetails INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_5 ON dbo.LotTransferDetails.VRD_From_RefID = VendRcvdDetail_5.EntryID LEFT OUTER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplit ON 
                      Tab2.LotNo = TSplit.LotNo AND Tab1.ProcessID = TSplit.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_4.LotNo AS FromLotNo, VendRcvdDetail_4.ProcessID, LotTransferDetails_1.SplitQty AS SplitQtyFrom, 
                                                   LotTransferDetails_1.Type AS SplitTypeFrom, LotTransferDetails_1.FromOrderNo, VendRcvdDetail_1.LotNo, 
                                                   LotTransferDetails_1.LotTransferRemarks
                             FROM         dbo.LotTransferDetails AS LotTransferDetails_1 INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_4 ON LotTransferDetails_1.VRD_From_RefID = VendRcvdDetail_4.EntryID INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON LotTransferDetails_1.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplitFrom ON 
                      Tab2.LotNo = TSplitFrom.LotNo AND Tab1.ProcessID = TSplitFrom.ProcessID LEFT OUTER JOIN
                          (SELECT     VendIssued_1.ProcessID, VendIssdDetail_1.LotNo, '' AS IssSFRemarks, VendIssdDetail_1.IssQty AS IssSFQty,TSOO_I_1.StoreName,TSOO_I_1.Location
                             FROM         dbo.VendIssdDetail AS VendIssdDetail_1 INNER JOIN
                                                       (SELECT     VID_RefID,StoreName,Location
                                                          FROM         dbo.StockOrderOpening_Issuance AS StockOrderOpening_Issuance_1
														  LEFT JOIN StockOrderOpening ON StockOrderOpening_Issuance_1.SOO_RefID=StockOrderOpening.EntryID
														  LEFT JOIN VStoreShelfs ON StockOrderOpening.Shelf_RefID=VStoreShelfs.EntryID
                                                          GROUP BY VID_RefID,StoreName,Location) AS TSOO_I_1 ON VendIssdDetail_1.EntryID = TSOO_I_1.VID_RefID INNER JOIN
                                                   dbo.VendIssued AS VendIssued_1 ON VendIssdDetail_1.RefID = VendIssued_1.EntryID) AS TIssFromSF ON Tab2.LotNo = TIssFromSF.LotNo AND 
                      Tab1.ProcessID = TIssFromSF.ProcessID LEFT OUTER JOIN
                          (SELECT     LotNo, MAX(EntryID) AS LotLastEntryID
                             FROM         dbo.VendRcvdDetail AS VendRcvdDetail_3 WHERE LotNo=@LotNo
                             GROUP BY LotNo) AS TLastRcv ON Tab2.LotNo = TLastRcv.LotNo LEFT OUTER JOIN
                      dbo.VLotWithPolishingMaker ON TSplitFrom.FromLotNo = dbo.VLotWithPolishingMaker.LotNo LEFT OUTER JOIN
                      dbo.VLotWithFirstProcessOfIssuance ON TSplitFrom.FromLotNo = dbo.VLotWithFirstProcessOfIssuance.LotNo
					  LEFT OUTER JOIN VLotsWithMasterPONo ON Tab2.LotNo=VLotsWithMasterPONo.LotNO
					  LEFT OUTER JOIN Lots_List ON Tab2.LotNo=Lots_List.LotNo
					  LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo, SUM(dbo.TransferredToReadyFinishLotsDetail.Qty) AS TTRFQty
                                FROM            dbo.TransferredToReadyFinishLotsDetail INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.TransferredToReadyFinishLotsDetail.VRD_RefID = dbo.VendRcvdDetail.EntryID
                                WHERE VendRcvdDetail.LotNo=@LotNo GROUP BY dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo ) AS Tab_TTRF ON Tab1.ProcessID = Tab_TTRF.ProcessID AND Tab2.LotNo = Tab_TTRF.LotNo
	WHERE Tab2.LotNo=@LotNo
	
	END
/*ELSE
	BEGIN
	
	SET @RepairGeneratedFrom=dbo.RepairLot_OriginalLots_F(@LotNo)
	SELECT        TFOrderItems.OrderNo, TFOrderItems.ItemCode, TFOrderItems.Qty, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.DT, 
                         dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.FCustomerOrders.Packaging, dbo.FCustomerOrders.DeliveryDT, 
                         dbo.RepairTypeProcesses.SeqNo AS SNO, dbo.RepairTypeProcesses.ProcessID, dbo.Processes.Description, Tab1.RcvdQty, Tab2.LotNo, Tab1.VendID, 
                         dbo.Makers.VenderName, Tab1.RcvDT, TabIss.IssDT, TabMin.MinProcessID, Tab1.Wastage, Tab1.PassQty, dbo.Employees.empid, dbo.Employees.name, 
                         Tab1.ReWorkQty, TFOrderItems.CompItemCode, dbo.Processes.Code, dbo.Processes.ProcessNameUrdu, dbo.Items.ItemNameUrdu, 
                         EmpRcv.name AS RcvEmpName, Tab1.Rate, TMasterPO.MasterPONo, TSteelBatch.SteelBatchNo, TForgeBatch.ForgeBatchNo, TSFRemarks.SFRemarks, 
                         TSplit.SplitQty, ISNULL(TSplit.Type, TSplitFrom.SplitTypeFrom) AS SplitType, TSplit.ToOrderNo, TSplit.ToLotNo, TSplitFrom.FromOrderNo, TSplitFrom.FromLotNo, 
                         TSplitFrom.SplitQtyFrom, @RepairGeneratedFrom AS RepairGeneratedFrom, dbo.RepairTypes.RepairType, TLastRcv.LotLastEntryID, 
                         Tab_TTRF.TTRFQty, Tab1.ReceivedEntryID, dbo.VendReceived_Employees_F(Tab1.ReceivedEntryID) AS Employees,'' AS IssEmpID,'' AS IssEmpName,dbo.GetMainLots_Repair_F(@LotNo) AS RepairLotsGenerated
FROM            (SELECT        OrderNo, ItemCode, CompItemCode, SUM(Qty) AS Qty
                           FROM            dbo.FOrderItems
                           GROUP BY OrderNo, ItemCode, CompItemCode) AS TFOrderItems INNER JOIN
                         dbo.Items ON TFOrderItems.CompItemCode = dbo.Items.ItemID INNER JOIN
                         dbo.FCustomerOrders ON TFOrderItems.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                             (SELECT DISTINCT OrderNo, ItemCode, LotNo, Repair_RefID
                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo) AS Tab2 ON TFOrderItems.OrderNo = Tab2.OrderNo AND TFOrderItems.CompItemCode = Tab2.ItemCode INNER JOIN
                         dbo.RepairTypeProcesses ON dbo.RepairTypeProcesses.Repair_RefID = Tab2.Repair_RefID INNER JOIN
                         dbo.Processes ON dbo.RepairTypeProcesses.ProcessID = dbo.Processes.ProcessID INNER JOIN
                         dbo.RepairTypes ON Tab2.Repair_RefID = dbo.RepairTypes.EntryID LEFT OUTER JOIN
                             (SELECT        VendRcvdDetail_2.ProcessID, VendRcvdDetail_2.LotNo, VendRcvdDetail_2.Rate, dbo.VendReceived.EmpID, VendRcvdDetail_2.RcvdQty, 
                                                         CASE WHEN ReqAuth = 1 THEN 0 ELSE RcvdQty END AS PassQty, dbo.VendReceived.VendID, dbo.VendReceived.DT AS RcvDT, 
                                                         VendRcvdDetail_2.Wastage, VendRcvdDetail_2.ReWorkQty, dbo.VendReceived.EntryID AS ReceivedEntryID
                                FROM            dbo.VendRcvdDetail AS VendRcvdDetail_2 INNER JOIN
                                                         dbo.VendReceived ON dbo.VendReceived.EntryID = VendRcvdDetail_2.RefID) AS Tab1 ON Tab2.LotNo = Tab1.LotNo AND 
                         dbo.RepairTypeProcesses.ProcessID = Tab1.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssdDetail.Repair_RefID, 
                                                         MIN(dbo.VendIssued.DT) AS IssDT
                                FROM            dbo.VendIssued INNER JOIN
                                                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID LEFT OUTER JOIN
                                                         dbo.VendIssdDetail_MoreDetails ON dbo.VendIssdDetail.EntryID = dbo.VendIssdDetail_MoreDetails.VID_RefID WHERE dbo.VendIssdDetail.LotNo=@LotNo
                                GROUP BY dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssdDetail.Repair_RefID) AS TabIss ON 
                         Tab2.LotNo = TabIss.LotNo AND dbo.RepairTypeProcesses.ProcessID = TabIss.ProcessID LEFT OUTER JOIN
                             (SELECT        VendRcvdDetail_1.LotNo, MIN(VendReceived_1.ProcessID) AS MinProcessID
                                FROM            dbo.VendReceived AS VendReceived_1 INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON VendReceived_1.EntryID = VendRcvdDetail_1.RefID WHERE VendRcvdDetail_1.LotNo=@LotNo
                                GROUP BY VendRcvdDetail_1.LotNo) AS TabMin ON Tab2.LotNo = TabMin.LotNo LEFT OUTER JOIN
                         dbo.Makers ON Tab1.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                         dbo.Employees ON TabIss.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                         dbo.Employees AS EmpRcv ON Tab1.EmpID = EmpRcv.empid LEFT OUTER JOIN
                             (SELECT        TNest1.LotNo, dbo.VendIssued.EntryID, dbo.VendIssued.MasterPONo
                                FROM            dbo.VendIssued INNER JOIN
                                                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
                                                             (SELECT        LotNo, MIN(Issue_RefID) AS Issue_RefID
                                                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo
                                                                GROUP BY LotNo) AS TNest1 ON TNest1.Issue_RefID = dbo.VendIssdDetail.EntryID) AS TMasterPO ON 
                         Tab1.LotNo = TMasterPO.LotNo LEFT OUTER JOIN
                             (SELECT        dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS SteelBatchNo
                                FROM            dbo.RawMaterialIssuance INNER JOIN
                                                         dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                         dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
                                WHERE        (dbo.RM.GroupID = 20)
                                GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TSteelBatch ON TMasterPO.EntryID = TSteelBatch.VI_RefID LEFT OUTER JOIN
                             (SELECT        dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS ForgeBatchNo
                                FROM            dbo.RawMaterialIssuance INNER JOIN
                                                         dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                         dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
                                WHERE        (dbo.RM.GroupID IN (25, 27, 28, 29, 30))
                                GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TForgeBatch ON TMasterPO.EntryID = TForgeBatch.VI_RefID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo, dbo.StockOrderOpening.Remarks AS SFRemarks
                                FROM            dbo.StockOrderOpening INNER JOIN
                                                         dbo.StockOrderOpening_Issuance ON dbo.StockOrderOpening.EntryID = dbo.StockOrderOpening_Issuance.SOO_RefID INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.StockOrderOpening_Issuance.VID_RefID = dbo.VendRcvdDetail.Issue_RefID) AS TSFRemarks ON 
                         Tab2.LotNo = TSFRemarks.LotNo AND Tab1.ProcessID = TSFRemarks.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.LotNo, dbo.VendRcvdDetail.ProcessID, dbo.LotTransferDetails.SplitQty, dbo.LotTransferDetails.Type, 
                                                         dbo.LotTransferDetails.ToOrderNo, VendRcvdDetail_1.LotNo AS ToLotNo
                                FROM            dbo.LotTransferDetails INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.LotTransferDetails.VRD_From_RefID = dbo.VendRcvdDetail.EntryID INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplit ON 
                         Tab2.LotNo = TSplit.LotNo AND Tab1.ProcessID = TSplit.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.LotNo AS FromLotNo, dbo.VendRcvdDetail.ProcessID, dbo.LotTransferDetails.SplitQty AS SplitQtyFrom, 
                                                         dbo.LotTransferDetails.Type AS SplitTypeFrom, dbo.LotTransferDetails.FromOrderNo, VendRcvdDetail_1.LotNo
                                FROM            dbo.LotTransferDetails INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.LotTransferDetails.VRD_From_RefID = dbo.VendRcvdDetail.EntryID INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplitFrom ON 
                         Tab2.LotNo = TSplitFrom.LotNo AND Tab1.ProcessID = TSplitFrom.ProcessID LEFT OUTER JOIN
                             (SELECT        LotNo, MAX(EntryID) AS LotLastEntryID
                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo
                                GROUP BY LotNo) AS TLastRcv ON Tab2.LotNo = TLastRcv.LotNo LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo, SUM(dbo.TransferredToReadyFinishLotsDetail.Qty) AS TTRFQty
                                FROM            dbo.TransferredToReadyFinishLotsDetail INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.TransferredToReadyFinishLotsDetail.VRD_RefID = dbo.VendRcvdDetail.EntryID
                                WHERE VendRcvdDetail.LotNo=@LotNo GROUP BY dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo ) AS Tab_TTRF ON Tab1.ProcessID = Tab_TTRF.ProcessID AND 
                         Tab2.LotNo = Tab_TTRF.LotNo
							WHERE Tab2.LotNo=@LotNo
		END	*/
--SET ANSI_WARNINGS ON
END


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Running_Lots_Summary_Hubwise]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[Running_Lots_Summary_Hubwise];
END
GO

-- ====================================================================================================
-- PART 6: STORED PROCEDURES
-- ====================================================================================================
CREATE   PROCEDURE Running_Lots_Summary_Hubwise(@DTFrom DATETIME,@DTTo DATETIME,@OrderNo VARCHAR(50)='0')

AS
BEGIN

	SET NOCOUNT ON;

    SELECT 
	COUNT(DISTINCT (CASE WHEN Hub_Name='H1' THEN ItemCode ELSE NULL END)) AS H1_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H2' THEN ItemCode ELSE NULL END)) AS H2_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H3' THEN ItemCode ELSE NULL END)) AS H3_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H4' THEN ItemCode ELSE NULL END)) AS H4_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H5' THEN ItemCode ELSE NULL END)) AS H5_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H6' THEN ItemCode ELSE NULL END)) AS H6_Items
	,SUM(CASE WHEN Hub_Name='H1' THEN Qty ELSE 0 END) AS H1_Qty
	,SUM(CASE WHEN Hub_Name='H2' THEN Qty ELSE 0 END) AS H2_Qty
	,SUM(CASE WHEN Hub_Name='H3' THEN Qty ELSE 0 END) AS H3_Qty
	,SUM(CASE WHEN Hub_Name='H4' THEN Qty ELSE 0 END) AS H4_Qty
	,SUM(CASE WHEN Hub_Name='H5' THEN Qty ELSE 0 END) AS H5_Qty
	,SUM(CASE WHEN Hub_Name='H6' THEN Qty ELSE 0 END) AS H6_Qty
	FROM VRunningLots
	INNER JOIN ItemProcessGroups ON VRunningLots.ItemCode=ItemProcessGroups.ItemID
	INNER JOIN ProcessGroupsProcesses ON ItemProcessGroups.PG_RefID=ProcessGroupsProcesses.Group_RefID AND VRunningLots.ProcessID=ProcessGroupsProcesses.Process_RefID
	WHERE VRunningLots.DT BETWEEN @DTFrom AND @DTTo
	AND (VRunningLots.OrderNo=@OrderNo OR @OrderNo IN('0',''))
	GROUP BY Hub_Name

END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_VendAssItem]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[SP_VendAssItem];
END
GO
CREATE   PROCEDURE [dbo].[SP_VendAssItem] 
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
        INSERT INTO VendAssItems(VendID, ProcessID, ItemID, Rate, PlantRate, StampRate, SnaffRate, Unit, CustCode)
        SELECT @VendID, @ProcessID, ItemID, @Rate, PlantRate, StampRate, SnaffRate, Unit, @EffectiveCustCode
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
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Track_Remaining_Qty_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[Track_Remaining_Qty_SP];
END
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Track_Remaining_Qty_SP] (@OrderNo VARCHAR(50))
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

SELECT FCO.CustCode,FCO.Country,FCO.OrderNO,FOI.CompItemCode,FOI.ItemCode,VItems.ItemName,FOI.Qty AS OrderQty
	,TSOA.SOA_Qty,TRcv.Received_Qty,TShipped.Shipped_Qty,FOI.Qty-ISNULL(TShipped.Shipped_Qty,0) AS Remaining_Qty
	,TMTS.Move_To_Stock_Qty
	,TRR.Rework_Qty,TRR.Rejection_Qty
	,FCO.DT AS OrderDT
	,VItems.InHand
	,TShipped.Shipped_Qty AS ItemShippedQty
	FROM FCustomerOrders FCO
	INNER JOIN FOrderItems FOI ON FCO.OrderNo=FOI.OrderNo
	INNER JOIN VItems ON FOI.CompItemCode=VItems.ItemID

	LEFT JOIN (SELECT OrderNo,ItemID,SUM(Qty) AS SOA_Qty FROM StockOrderAdjustments WHERE OrderNo=@OrderNo GROUP BY OrderNo,ItemID) TSOA ON FOI.OrderNo=TSOA.OrderNo AND FOI.CompItemCode=TSOA.ItemID
	LEFT JOIN (SELECT OrderNo,ItemCode,SUM(RcvdQty) AS Received_Qty FROM VendRcvdDetail 
					INNER JOIN VendReceived ON VendReceived.EntryID=VendRcvdDetail.RefID 
					INNER JOIN VendIssued ON VendReceived.Issuance_RefID=VendIssued.EntryID
					WHERE OrderNo=@OrderNo AND MasterPONo IS NULL GROUP BY OrderNo,ItemCode) TRcv ON FOI.OrderNo=TRcv.OrderNo AND FOI.CompItemCode=TRcv.ItemCode

	LEFT JOIN (SELECT OrderNo,ItemID,SUM(Qty) AS Shipped_Qty FROM VCustomInvoiceItems WHERE OrderNo=@OrderNO GROUP BY OrderNo,ItemID) TShipped ON FOI.OrderNo=TShipped.OrderNo AND FOI.CompItemCode=TShipped.ItemID
	LEFT JOIN (SELECT OrderNo,ItemCode,SUM(TTRFLD.Qty) AS Move_To_Stock_Qty FROM TransferredToReadyFinishLotsDetail TTRFLD
					INNER JOIN vendRcvdDetail VRD ON TTRFLD.VRD_RefID=VRD.EntryID WHERE OrderNo=@OrderNo GROUP BY OrderNo,ItemCode) TMTS ON FOI.OrderNo=TMTS.OrderNo AND FOI.CompItemCode=TMTS.ItemCode
	LEFT JOIN (SELECT OrderNo,ItemCode,SUM(ReWorkQty) AS Rework_Qty,SUM(Wastage) AS Rejection_Qty FROM VendRcvdDetail WHERE OrderNo=@OrderNo GROUP BY OrderNo,ItemCode) TRR ON FOI.OrderNo=TRR.OrderNo AND FOI.CompItemCode=TRR.ItemCode
	WHERE FCO.OrderNo=@OrderNo


    
END

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Trial_Balance_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[Trial_Balance_SP];
END
GO

-- =============================================
-- Author:		Asim Arif
-- Create date: Feb-2017
-- Description:	Trial Balance (Fastest, Simplest)
-- =============================================
CREATE   PROCEDURE [dbo].[Trial_Balance_SP] (@DTFrom DATETIME,@DTTo DATETIME,@Code VARCHAR(50)='0')
	
AS
BEGIN
	
	SET NOCOUNT ON;
	
	DECLARE @CodeLength AS INT
	IF ISNULL(@Code,'0')='0'
		SET @CodeLength=0
	ELSE
		SET @CodeLength=LEN(@Code)

	CREATE TABLE #Accs(AccNo VARCHAR(50),Opening_Balance BIGINT,Opening_Debit BIGINT,Opening_Credit BIGINT,Debit BIGINT,Credit BIGINT,Closing_Balance BIGINT,Balance_Debit BIGINT,Balance_Credit BIGINT,EntryType TINYINT,AccType VARCHAR(50))
	CREATE TABLE #Accs_Parent(AccNo VARCHAR(50),Opening_Balance BIGINT,Opening_Debit BIGINT,Opening_Credit BIGINT,Debit BIGINT,Credit BIGINT,Closing_Balance BIGINT,Balance_Debit BIGINT,Balance_Credit BIGINT,EntryType TINYINT,AccType VARCHAR(50))
	CREATE TABLE #Accs_Heads(AccNo VARCHAR(50),Opening_Balance BIGINT,Opening_Debit BIGINT,Opening_Credit BIGINT,Debit BIGINT,Credit BIGINT,Closing_Balance BIGINT,Balance_Debit BIGINT,Balance_Credit BIGINT,EntryType TINYINT,AccType VARCHAR(50))

	INSERT INTO #Accs(AccNo,Opening_Balance,Closing_Balance,Debit,Credit,EntryType,AccType)
		SELECT Accounts.AccNo,ISNULL(TOpening.Balance,Accounts.openbal),ISNULL(TClosing.Balance,Accounts.openbal),TTotal.TotalDebit,TTotal.TotalCredit,0,Accounts.Type FROM Accounts
			LEFT JOIN (SELECT AccNo,Balance FROM Vouchers INNER JOIN (SELECT MAX(SNo) AS MaxSNo FROM Vouchers WHERE VDate<@DTFrom GROUP BY AccNo) T1 ON Vouchers.SNo=T1.MaxSNo) TOpening ON Accounts.AccNo=TOpening.Accno
			LEFT JOIN (SELECT AccNo,Balance FROM Vouchers INNER JOIN (SELECT MAX(SNo) AS MaxSNo FROM Vouchers WHERE VDate<=@DTTo GROUP BY AccNo) T1 ON Vouchers.SNo=T1.MaxSNo) TClosing ON Accounts.AccNo=TClosing.Accno
			LEFT JOIN (SELECT AccNo,SUM(Debit) AS TotalDebit,SUM(Credit) AS TotalCredit FROM Vouchers WHERE VDate BETWEEN @DTFrom AND @DTTo GROUP BY Accno) TTotal ON Accounts.AccNo=TTotal.AccNo
			WHERE Parent=0
			AND (LEFT(Accounts.AccNo,@CodeLength)=@Code OR ISNULL(@Code,'0')='0')
			--AND (Accounts.Account_Type=@Account_Type OR @Account_Type=-1)

	--INSERT INTO #Accs(AccNo,Opening_Debit,Opening_Credit,Debit,Credit,Balance_Debit,Balance_Credit)
	UPDATE #Accs SET	Opening_Debit=CASE WHEN Opening_Balance>=0 THEN Opening_Balance ELSE 0 END,
						Opening_Credit=CASE WHEN Opening_Balance<0 THEN ABS(Opening_Balance) ELSE 0 END,
						Balance_Debit=CASE WHEN Closing_Balance>=0 THEN Closing_Balance ELSE 0 END,
						Balance_Credit=CASE WHEN Closing_Balance<0 THEN ABS(Closing_Balance) ELSE 0 END

	INSERT INTO #Accs_Parent(AccNo,EntryType,AccType)
		SELECT AccNo,1,Accounts.Type FROM Accounts
			WHERE Parent=1
			AND (LEFT(Accounts.AccNo,@CodeLength)=@Code OR ISNULL(@Code,'0')='0')

	UPDATE T1 SET	T1.Opening_Debit=(SELECT SUM(Opening_Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Opening_Credit=(SELECT SUM(Opening_Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Balance_Debit=(SELECT SUM(Balance_Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Balance_Credit=(SELECT SUM(Balance_Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Debit=(SELECT SUM(Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Credit=(SELECT SUM(Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo)
		FROM #Accs_Parent T1

	INSERT INTO #Accs_Heads(AccNo,EntryType,AccType)
		SELECT Code,2,Code FROM Heads
		
	UPDATE T1 SET	T1.Opening_Debit=(SELECT SUM(Opening_Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Opening_Credit=(SELECT SUM(Opening_Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Balance_Debit=(SELECT SUM(Balance_Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Balance_Credit=(SELECT SUM(Balance_Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Debit=(SELECT SUM(Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Credit=(SELECT SUM(Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo)
		FROM #Accs_Heads T1
		
		
	SELECT * FROM 
	(SELECT #Accs.*,Accounts.AccTitle FROM #Accs INNER JOIN Accounts ON #Accs.AccNo=Accounts.AccNo
	UNION ALL
	SELECT #Accs_Parent.*,Accounts.AccTitle FROM #Accs_Parent INNER JOIN Accounts ON #Accs_Parent.AccNo=Accounts.AccNo
	UNION ALL
	SELECT #Accs_Heads.*,Heads.Head FROM #Accs_Heads INNER JOIN Heads ON #Accs_Heads.AccNo=Heads.Code) TUnioned 
	--WHERE (AccType=@Code OR @Code='0')
	--WHERE (LEFT(TUnioned.AccNo,@CodeLength)=@Code OR @Code IS NULL)

	ORDER BY AccNo
	
	DROP TABLE #Accs
	DROP TABLE #Accs_Parent
	DROP TABLE #Accs_Heads
END


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Trial_Balance_SP_Ex]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[Trial_Balance_SP_Ex];
END
GO
-- =============================================
-- Author:		Asim Arif
-- Refactored:  Performance Optimized (No Resultset Changes)
-- Create date: Feb-2017
-- Description:	Trial Balance (Fastest, Simplest)
-- =============================================
CREATE   PROCEDURE [dbo].[Trial_Balance_SP_Ex] (
    @DTFrom DATETIME,
    @DTTo DATETIME,
    @Code VARCHAR(50)='0'
)
AS
BEGIN
    SET NOCOUNT ON;
	
    DECLARE @CodeLength AS INT
    IF ISNULL(@Code,'0') = '0'
        SET @CodeLength = 0
    ELSE
        SET @CodeLength = LEN(@Code)

    -- Single temp table for base accounts (Parent = 0)
    CREATE TABLE #Accs (
        AccNo VARCHAR(50),
        Opening_Balance BIGINT,
        Opening_Debit BIGINT,
        Opening_Credit BIGINT,
        Debit BIGINT,
        Credit BIGINT,
        Closing_Balance BIGINT,
        Balance_Debit BIGINT,
        Balance_Credit BIGINT,
        EntryType TINYINT,
        AccType VARCHAR(50)
    )

    -- 1. Insert and calculate child accounts directly without UPDATE passes
    INSERT INTO #Accs (
        AccNo, Opening_Balance, Closing_Balance,
        Opening_Debit, Opening_Credit,
        Balance_Debit, Balance_Credit,
        Debit, Credit, EntryType, AccType
    )
    SELECT 
        A.AccNo, 
        Base.OpBal, 
        Base.ClBal,
        CASE WHEN Base.OpBal >= 0 THEN Base.OpBal ELSE 0 END,
        CASE WHEN Base.OpBal < 0 THEN ABS(Base.OpBal) ELSE 0 END,
        CASE WHEN Base.ClBal >= 0 THEN Base.ClBal ELSE 0 END,
        CASE WHEN Base.ClBal < 0 THEN ABS(Base.ClBal) ELSE 0 END,
        Tot.TotalDebit,
        Tot.TotalCredit,
        0,
        A.Type
    FROM Accounts A
    CROSS APPLY (
        -- Replaces expensive MAX(SNo) groupings with highly targeted Top 1 index seeks
        SELECT 
            OpBal = ISNULL((SELECT TOP 1 Balance FROM Vouchers V WHERE V.AccNo = A.AccNo AND V.VDate < @DTFrom ORDER BY SNo DESC), A.openbal),
            ClBal = ISNULL((SELECT TOP 1 Balance FROM Vouchers V WHERE V.AccNo = A.AccNo AND V.VDate <= @DTTo ORDER BY SNo DESC), A.openbal)
    ) Base
    OUTER APPLY (
        SELECT SUM(Debit) AS TotalDebit, SUM(Credit) AS TotalCredit 
        FROM Vouchers V 
        WHERE V.AccNo = A.AccNo AND V.VDate BETWEEN @DTFrom AND @DTTo
    ) Tot
    WHERE A.Parent = 0
      AND (LEFT(A.AccNo, @CodeLength) = @Code OR @CodeLength = 0)


    -- 2. Output combined resultset directly (Base + Parents + Heads)
    SELECT 
        A.AccNo, A.Opening_Balance, A.Opening_Debit, A.Opening_Credit, 
        A.Debit, A.Credit, A.Closing_Balance, A.Balance_Debit, A.Balance_Credit, 
        A.EntryType, A.AccType, Acc.AccTitle
    FROM #Accs A
    INNER JOIN Accounts Acc ON A.AccNo = Acc.AccNo

    UNION ALL

    -- Aggregate Parent Accounts (Parent = 1)
    SELECT 
        P.AccNo, 
        CAST(NULL AS BIGINT) AS Opening_Balance,
        SUM(C.Opening_Debit) AS Opening_Debit,
        SUM(C.Opening_Credit) AS Opening_Credit,
        SUM(C.Debit) AS Debit,
        SUM(C.Credit) AS Credit,
        CAST(NULL AS BIGINT) AS Closing_Balance,
        SUM(C.Balance_Debit) AS Balance_Debit,
        SUM(C.Balance_Credit) AS Balance_Credit,
        CAST(1 AS TINYINT) AS EntryType, 
        P.Type AS AccType,
        P.AccTitle
    FROM Accounts P
    LEFT JOIN #Accs C ON LEFT(C.AccNo, LEN(P.AccNo)) = P.AccNo
    WHERE P.Parent = 1
      AND (LEFT(P.AccNo, @CodeLength) = @Code OR @CodeLength = 0)
    GROUP BY P.AccNo, P.Type, P.AccTitle

    UNION ALL

    -- Aggregate Head Accounts
    SELECT 
        H.Code AS AccNo, 
        CAST(NULL AS BIGINT) AS Opening_Balance,
        SUM(C.Opening_Debit) AS Opening_Debit,
        SUM(C.Opening_Credit) AS Opening_Credit,
        SUM(C.Debit) AS Debit,
        SUM(C.Credit) AS Credit,
        CAST(NULL AS BIGINT) AS Closing_Balance,
        SUM(C.Balance_Debit) AS Balance_Debit,
        SUM(C.Balance_Credit) AS Balance_Credit,
        CAST(2 AS TINYINT) AS EntryType, 
        H.Code AS AccType,
        H.Head AS AccTitle
    FROM Heads H
    LEFT JOIN #Accs C ON LEFT(C.AccNo, 2) = H.Code
    GROUP BY H.Code, H.Head

    ORDER BY AccNo
	
    DROP TABLE #Accs
END


GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[User_Activity_Report_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[User_Activity_Report_SP];
END
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE   PROCEDURE [dbo].[User_Activity_Report_SP] (@UserName VARCHAR(50),@DTFrom DATETIME,@DTTO DATETIME)

AS
BEGIN

	SET @DTFrom=CAST(CONVERT(VARCHAR(50),@DTFrom,6) AS DATETIME)
	SET @DTTo=CAST(CONVERT(VARCHAR(50),@DTTo,6) AS DATETIME)
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	SELECT T1.UserName,T1.DT,T1.ItemCode,T1.Qty,T1.ProcessID,T1.EntryType,Items.ItemName,
		   Processes.Description AS Process_Description,DTEntry,MachineName
	 FROM (
		SELECT VI.UserName,VI.DT,VID.ItemCode,VID.IssQty AS Qty,VI.ProcessID,
		CASE WHEN MasterPONo IS NULL THEN 'Process Issuance' ELSE 'Purcahse Order' END AS EntryType ,
		VI.DTEntry,VI.MachineName
		FROM VendIssued VI
		INNER JOIN VendIssdDetail VID ON VI.EntryID=VID.RefID 	
		INNER JOIN Makers ON VI.VendID=Makers.VendID
		WHERE VI.DT BETWEEN @DTFrom AND @DTTo
		AND VI.UserName=@UserName

	UNION ALL


	SELECT VR.UserName,VR.DT,VRD.ItemCode,VRD.RcvdQty,VRD.ProcessID,CASE WHEN VI.MasterPONo IS NULL THEN 'Process Receiving' ELSE 'Receiving' END AS EntryType,
		 VR.EntryDT,VR.MachineName
	FROM VendReceived VR 
		INNER JOIN VendRcvdDetail VRD ON VR.EntryID=VRD.RefID
		INNER JOIN VendIssued VI ON VR.Issuance_RefID=VI.EntryID
		WHERE VR.DT BETWEEN @DTFrom AND @DTTo
		AND VR.UserName=@UserName
		) T1

		INNER JOIN Items ON T1.ItemCode=Items.ItemID
		INNER JOIN Processes ON T1.ProcessID=Processes.ProcessID
		
END


SELECT * FROM VendIssued

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VendorLoanBalanceReport_SP]') AND type IN (N'P'))
BEGIN
    DROP PROCEDURE [dbo].[VendorLoanBalanceReport_SP];
END
GO

CREATE   PROCEDURE [dbo].[VendorLoanBalanceReport_SP] (@DT AS DATETIME,@InActive AS BIT) AS
SET @DT=CAST(CONVERT(VARCHAR(50),@DT,6) AS DATETIME)

SELECT     TotalAmountTaken-TotalAmountCleared AS LongTermBalance, VVendersAndMakers.SubAccOf, VVendersAndMakers.AccTitle_Parent, VVendersAndMakers.AccNo, VVendersAndMakers.AccTitle,VVendersAndMakers.MakerNo,
                      Amount-AmountCleared AS STLoan,ActualDeductionAmount AS LTLoanDed,Active
FROM         VVendersAndMakers
		LEFT OUTER JOIN VMakerAdvancesShortDeduction ON VVendersANDMakers.AccNo=VMakerAdvancesShortDeduction.AccNo
		LEFT OUTER JOIN VMakerAdvancesDeduction ON VVendersAndMakers.AccNo=VMakerAdvancesDeduction.AccNo
	WHERE (Active=0 OR @InActive=0)

GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

-- =============================================================================
-- PHASE 7: REFRESH ALL VIEWS
-- =============================================================================
DECLARE @vName nvarchar(512);
DECLARE v_cur CURSOR LOCAL FAST_FORWARD FOR
    SELECT QUOTENAME(s.name) + '.' + QUOTENAME(v.name)
    FROM sys.views v
    JOIN sys.schemas s ON v.schema_id = s.schema_id
    WHERE v.is_ms_shipped = 0
    ORDER BY v.name;
OPEN v_cur;
FETCH NEXT FROM v_cur INTO @vName;
WHILE @@FETCH_STATUS = 0
BEGIN
    BEGIN TRY
        EXEC sp_refreshview @vName;
    END TRY
    BEGIN CATCH
        PRINT 'Warning refreshing ' + @vName + ': ' + ERROR_MESSAGE();
    END CATCH
    FETCH NEXT FROM v_cur INTO @vName;
END
CLOSE v_cur;
DEALLOCATE v_cur;
GO
