-- ==============================================================================
-- MIGRATION: ADD BUDGETING, FIXED ASSETS, AND GATE MANAGEMENT MODULES
-- Generated for seamless multi-module integration without breaking existing schema
-- ==============================================================================

-- ==============================================================================
-- PHASE 1: BUDGETING, FIXED ASSETS, AND GATE MANAGEMENT MIGRATION SCRIPT
-- Strictly Idempotent, Safe, and Database-Independent (No DROP, No destructive ALTER)
-- ==============================================================================

-- 1. COLUMNS ADDITIONS TO EXISTING TABLES
-- ------------------------------------------------------------------------------

-- Table: dbo.Accounts
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Accounts') AND name = 'BudgetAccount')
    ALTER TABLE dbo.Accounts ADD [BudgetAccount] bit NOT NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Accounts') AND name = 'BudgetAccount_Glance')
    ALTER TABLE dbo.Accounts ADD [BudgetAccount_Glance] bit NOT NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Accounts') AND name = 'IsNonTransaction')
    ALTER TABLE dbo.Accounts ADD [IsNonTransaction] bit NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Accounts') AND name = 'ShowInExpenseReport')
    ALTER TABLE dbo.Accounts ADD [ShowInExpenseReport] bit NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Accounts') AND name = 'ShowInBalanceSheet')
    ALTER TABLE dbo.Accounts ADD [ShowInBalanceSheet] bit NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Accounts') AND name = 'SpecialAccount')
    ALTER TABLE dbo.Accounts ADD [SpecialAccount] bit NULL DEFAULT ((0));
GO

-- Table: dbo.Heads
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Heads') AND name = 'BudgetAccount_Glance')
    ALTER TABLE dbo.Heads ADD [BudgetAccount_Glance] bit NOT NULL DEFAULT ((0));
GO

-- Table: dbo.FixedAssets
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'AcumulatedDepriciation')
    ALTER TABLE dbo.FixedAssets ADD [AcumulatedDepriciation] real NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'DepriciationRate')
    ALTER TABLE dbo.FixedAssets ADD [DepriciationRate] real NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'VRDPO_RefID')
    ALTER TABLE dbo.FixedAssets ADD [VRDPO_RefID] int NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'OldCode')
    ALTER TABLE dbo.FixedAssets ADD [OldCode] varchar(255) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'Disposed')
    ALTER TABLE dbo.FixedAssets ADD [Disposed] bit NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'DTDisposal')
    ALTER TABLE dbo.FixedAssets ADD [DTDisposal] datetime NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'SalePrice')
    ALTER TABLE dbo.FixedAssets ADD [SalePrice] real NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'WDV')
    ALTER TABLE dbo.FixedAssets ADD [WDV] real NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'Remarks')
    ALTER TABLE dbo.FixedAssets ADD [Remarks] varchar(1000) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'From_JV')
    ALTER TABLE dbo.FixedAssets ADD [From_JV] bit NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'VchrNo')
    ALTER TABLE dbo.FixedAssets ADD [VchrNo] varchar(50) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'Disposed_Type')
    ALTER TABLE dbo.FixedAssets ADD [Disposed_Type] tinyint NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'EmpID')
    ALTER TABLE dbo.FixedAssets ADD [EmpID] varchar(50) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'Approved')
    ALTER TABLE dbo.FixedAssets ADD [Approved] tinyint NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'RcvProcessDT')
    ALTER TABLE dbo.FixedAssets ADD [RcvProcessDT] datetime NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'UserName')
    ALTER TABLE dbo.FixedAssets ADD [UserName] varchar(50) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'MachineName')
    ALTER TABLE dbo.FixedAssets ADD [MachineName] varchar(50) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'CaseClosed')
    ALTER TABLE dbo.FixedAssets ADD [CaseClosed] tinyint NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'CC_UserName')
    ALTER TABLE dbo.FixedAssets ADD [CC_UserName] varchar(50) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'CC_MachineName')
    ALTER TABLE dbo.FixedAssets ADD [CC_MachineName] varchar(50) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'CC_EntryDT')
    ALTER TABLE dbo.FixedAssets ADD [CC_EntryDT] datetime NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'EntryDT')
    ALTER TABLE dbo.FixedAssets ADD [EntryDT] datetime NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets') AND name = 'InActiveDT')
    ALTER TABLE dbo.FixedAssets ADD [InActiveDT] datetime NULL;
GO

-- Table: dbo.FixedAssets_Master
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets_Master') AND name = 'PercentRate')
    ALTER TABLE dbo.FixedAssets_Master ADD [PercentRate] real NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets_Master') AND name = 'AccumAccNo')
    ALTER TABLE dbo.FixedAssets_Master ADD [AccumAccNo] nvarchar(510) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets_Master') AND name = 'DepriciationAccNo')
    ALTER TABLE dbo.FixedAssets_Master ADD [DepriciationAccNo] nvarchar(510) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets_Master') AND name = 'DepriciationMethodID')
    ALTER TABLE dbo.FixedAssets_Master ADD [DepriciationMethodID] tinyint NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets_Master') AND name = 'Tag')
    ALTER TABLE dbo.FixedAssets_Master ADD [Tag] tinyint NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets_Master') AND name = 'COA_AccNo')
    ALTER TABLE dbo.FixedAssets_Master ADD [COA_AccNo] varchar(50) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets_Master') AND name = 'DepAccNo_Dr')
    ALTER TABLE dbo.FixedAssets_Master ADD [DepAccNo_Dr] nvarchar(100) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.FixedAssets_Master') AND name = 'DepAccNo_Cr')
    ALTER TABLE dbo.FixedAssets_Master ADD [DepAccNo_Cr] nvarchar(100) NULL;
GO

-- Table: dbo.Users (Module Access & Gate Rights)
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'BudgetingMainLink')
    ALTER TABLE dbo.Users ADD [BudgetingMainLink] bit NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'FixedAssetsMainLink')
    ALTER TABLE dbo.Users ADD [FixedAssetsMainLink] bit NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'GateMainLink')
    ALTER TABLE dbo.Users ADD [GateMainLink] bit NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'AWMGateDashboard')
    ALTER TABLE dbo.Users ADD [AWMGateDashboard] bit NOT NULL DEFAULT ((1));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'AWMGateIn')
    ALTER TABLE dbo.Users ADD [AWMGateIn] bit NOT NULL DEFAULT ((1));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'AWMGatePassTypes')
    ALTER TABLE dbo.Users ADD [AWMGatePassTypes] bit NOT NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'AWMVisitors')
    ALTER TABLE dbo.Users ADD [AWMVisitors] bit NOT NULL DEFAULT ((1));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'AWMUserRights')
    ALTER TABLE dbo.Users ADD [AWMUserRights] bit NOT NULL DEFAULT ((0));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'AWMVehicleManagement')
    ALTER TABLE dbo.Users ADD [AWMVehicleManagement] bit NOT NULL DEFAULT ((1));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'AWMMaterialInward')
    ALTER TABLE dbo.Users ADD [AWMMaterialInward] bit NOT NULL DEFAULT ((1));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'AWMMaterialOutward')
    ALTER TABLE dbo.Users ADD [AWMMaterialOutward] bit NOT NULL DEFAULT ((1));
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'GateLastLoginAt')
    ALTER TABLE dbo.Users ADD [GateLastLoginAt] datetime NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'GateLastLogoutAt')
    ALTER TABLE dbo.Users ADD [GateLastLogoutAt] datetime NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Users') AND name = 'GateLastSeenAt')
    ALTER TABLE dbo.Users ADD [GateLastSeenAt] datetime NULL;
GO

-- ------------------------------------------------------------------------------
-- 2. NEW TABLES
-- ------------------------------------------------------------------------------

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'BudgetingAccounts')
BEGIN
    CREATE TABLE [dbo].[BudgetingAccounts] (
    [AccNo] nvarchar(255) NOT NULL,
    [AccTitle] nvarchar(255) NULL,
    [SubAccOf] nvarchar(50) NULL,
    [Type] nvarchar(50) NULL,
    [opendate] nvarchar(50) NULL,
    [openbal] money NULL,
    [Balance] money NULL,
    [Parent] bit NULL DEFAULT ((0)),
    [Active] bit NULL DEFAULT ((0)),
    [CreateDT] datetime NULL DEFAULT (getdate()),
    [UserName] varchar(50) NULL,
    [ComputerName] varchar(50) NULL,
    [Revision] real NULL,
    CONSTRAINT [PK_BudgetingAccounts] PRIMARY KEY CLUSTERED ([AccNo])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FinancialBudgeting')
BEGIN
    CREATE TABLE [dbo].[FinancialBudgeting] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [FinancialYear] varchar(50) NOT NULL,
    [AccNo] nvarchar(255) NOT NULL,
    [Month_July] bigint NULL,
    [Month_August] bigint NULL,
    [Month_September] bigint NULL,
    [Month_October] bigint NULL,
    [Month_November] bigint NULL,
    [Month_December] bigint NULL,
    [Month_January] bigint NULL,
    [Month_February] bigint NULL,
    [Month_March] bigint NULL,
    [Month_April] bigint NULL,
    [Month_May] bigint NULL,
    [Month_June] bigint NULL,
    [Jul_Sep] bigint NULL,
    [Jul_Sep_Revision] bigint NULL,
    [Oct_Dec] bigint NULL,
    [Oct_Dec_Revision] bigint NULL,
    [Jan_Mar] bigint NULL,
    [Jan_Mar_Revision] bigint NULL,
    [Apr_Jun] bigint NULL,
    [Apr_Jun_Revision] bigint NULL,
    [Month_July_R] bigint NULL,
    [Month_August_R] bigint NULL,
    [Month_September_R] bigint NULL,
    [Month_October_R] bigint NULL,
    [Month_November_R] bigint NULL,
    [Month_December_R] bigint NULL,
    [Month_January_R] bigint NULL,
    [Month_February_R] bigint NULL,
    [Month_March_R] bigint NULL,
    [Month_April_R] bigint NULL,
    [Month_May_R] bigint NULL,
    [Month_June_R] bigint NULL,
    CONSTRAINT [PK_FinancialBudgeting] PRIMARY KEY CLUSTERED ([FinancialYear], [AccNo])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FinancialBudgeting_Expenses')
BEGIN
    CREATE TABLE [dbo].[FinancialBudgeting_Expenses] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [FinancialYear] varchar(50) NOT NULL,
    [AccNo] nvarchar(255) NOT NULL,
    [Proposed] bigint NULL,
    [Revised] bigint NULL,
    CONSTRAINT [PK_FinancialBudgeting_Expenses] PRIMARY KEY CLUSTERED ([FinancialYear], [AccNo])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Adjustment')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Adjustment] (
    [EntryID] int NOT NULL,
    [AssetID] varchar(50) NULL,
    [DT] datetime NULL,
    [Remarks] varchar(500) NULL,
    [Cost] float NULL,
    [Adjustment] float NULL,
    [Attachment_Type] varchar(50) NULL,
    [PDFAttachments] varchar(255) NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Approval')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Approval] (
    [EntryID] int NOT NULL,
    [AssetID] varchar(50) NULL,
    [ForwardBy] varchar(50) NULL,
    [DTEntry] datetime NULL,
    [MachineName] varchar(50) NULL,
    [ForwardTo] varchar(50) NULL,
    [DTEntry_Done] datetime NULL,
    [MachineName_Done] varchar(50) NULL,
    [Remarks] varchar(5000) NULL,
    [Approval] int NULL,
    [ForwardBy_Designation] varchar(255) NULL,
    [ForwardTo_Designation] varchar(255) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Depreciation_Detail')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Depreciation_Detail] (
    [EntryID] int NOT NULL,
    [FA_D_M_RefID] int NULL,
    [AssetID] varchar(50) NULL,
    [Cost] float NULL,
    [BookValue] float NULL,
    [DisposedValue] real NULL,
    [OpeningBalance] float NULL,
    [AccountingYear] float NULL,
    [AccumulatedBalance] float NULL,
    [WrittenDownValue] float NULL,
    [Rate] real NULL,
    [Cost_Last] real NULL,
    [DisposedValue_Last] real NULL,
    [AcquisitionCost_Last] real NULL,
    [Adjustment_Opening] real NULL,
    [AdjustedValue] real NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Depreciation_Detail_Monthwise')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Depreciation_Detail_Monthwise] (
    [EntryID] int NOT NULL,
    [FA_D_M_RefID] int NULL,
    [AssetID] varchar(50) NULL,
    [Cost] real NULL,
    [BookValue] real NULL,
    [DisposedValue] real NULL,
    [OpeningBalance] real NULL,
    [AccountingYear] real NULL,
    [AccumulatedBalance] real NULL,
    [WrittenDownValue] real NULL,
    [Rate] real NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Depreciation_Main')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Depreciation_Main] (
    [EntryID] int NOT NULL,
    [DTFrom] datetime NULL,
    [DTTo] datetime NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Depreciation_Main_Monthwise')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Depreciation_Main_Monthwise] (
    [EntryID] int NOT NULL,
    [DTFrom] datetime NULL,
    [DTTo] datetime NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Disposals')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Disposals] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [AssetID] varchar(50) NOT NULL,
    [DisposalDate] datetime NOT NULL,
    [DisposalType] varchar(50) NULL,
    [DisposalAmount] decimal(18,2) NOT NULL DEFAULT ((0)),
    [Buyer] varchar(255) NULL,
    [Remarks] varchar(MAX) NULL,
    [UserName] varchar(100) NULL,
    [MachineName] varchar(100) NULL,
    [DTEntry] datetime NOT NULL DEFAULT (getdate()),
    CONSTRAINT [PK_FixedAssets_Disposals] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Head_Changes')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Head_Changes] (
    [EntryID] int NOT NULL,
    [AssetID] varchar(50) NULL,
    [FA_M_RefID_Old] int NULL,
    [FA_M_RefID_New] int NULL,
    [Remarks] varchar(255) NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Master_DepreciationRate')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Master_DepreciationRate] (
    [EntryID] int NOT NULL,
    [FA_M_RefID] int NULL,
    [DTFrom] datetime NULL,
    [DTTo] datetime NULL,
    [PercentRate] real NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FixedAssets_Master_Groups')
BEGIN
    CREATE TABLE [dbo].[FixedAssets_Master_Groups] (
    [Group_Name] varchar(255) NOT NULL,
    [FA_M_RefID] int NOT NULL,
    [Group_Name_Main] varchar(50) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'GateConcernedContacts')
BEGIN
    CREATE TABLE [dbo].[GateConcernedContacts] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [EmpID] varchar(50) NOT NULL,
    [ContactName] varchar(150) NULL,
    [EmailAddress] varchar(150) NULL,
    [WhatsAppNo] varchar(50) NULL,
    [Active] bit NOT NULL DEFAULT ((1)),
    [LastUpdatedBy] varchar(50) NULL,
    [LastUpdatedAt] datetime NOT NULL DEFAULT (getdate()),
    CONSTRAINT [PK_GateConcernedContacts] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'GateNotificationLog')
BEGIN
    CREATE TABLE [dbo].[GateNotificationLog] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [SourceModule] varchar(50) NOT NULL,
    [SourceEntryId] int NOT NULL,
    [ChannelName] varchar(20) NOT NULL,
    [RecipientName] varchar(150) NULL,
    [RecipientAddress] varchar(150) NULL,
    [MessageBody] varchar(MAX) NULL,
    [DeliveryStatus] varchar(30) NOT NULL DEFAULT ('Queued'),
    [DeliveryResponse] varchar(500) NULL,
    [CreatedAt] datetime NOT NULL DEFAULT (getdate()),
    [SentAt] datetime NULL,
    [CreatedBy] varchar(50) NULL,
    CONSTRAINT [PK_GateNotificationLog] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'GatePassTypes')
BEGIN
    CREATE TABLE [dbo].[GatePassTypes] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [GatePassType] varchar(50) NOT NULL,
    [Active] bit NOT NULL DEFAULT ((1)),
    CONSTRAINT [PK_GatePassTypes] PRIMARY KEY CLUSTERED ([GatePassType])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'GatePass_Return')
BEGIN
    CREATE TABLE [dbo].[GatePass_Return] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [GatePassNo] varchar(50) NULL,
    [DT] datetime NULL,
    [SenderName] varchar(255) NULL,
    [PersonName] varchar(255) NULL,
    [VehicleNo] varchar(50) NULL,
    [Description] varchar(255) NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL DEFAULT (getdate()),
    [Returnable] bit NULL DEFAULT ((0)),
    [IssdTo] varchar(255) NULL,
    [GatePassType] varchar(50) NULL,
    [ReturnDT] datetime NULL,
    [DeptID_From] varchar(50) NULL,
    [DeptID_To] varchar(50) NULL,
    [ApprovedBy_Name] varchar(50) NULL,
    [ApprovedBy_Designation] varchar(50) NULL,
    [VendID] int NULL,
    CONSTRAINT [PK_GatePass_Return] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'GatePass_Return_Details')
BEGIN
    CREATE TABLE [dbo].[GatePass_Return_Details] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [MEGO_RefID] int NULL,
    [Detail] varchar(255) NULL,
    [Unit] varchar(50) NULL,
    [Packing] varchar(255) NULL,
    [Qty] int NULL,
    CONSTRAINT [PK_GatePass_Return_Details] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'GatePass_Return_Receiving')
BEGIN
    CREATE TABLE [dbo].[GatePass_Return_Receiving] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL DEFAULT (getdate()),
    [ME_GO_RefID] int NULL,
    [RcvDT] datetime NULL,
    CONSTRAINT [PK_GatePass_Return_Receiving] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'GatePass_Return_Receiving_Detail')
BEGIN
    CREATE TABLE [dbo].[GatePass_Return_Receiving_Detail] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [ME_GO_Rcv_RefID] int NULL,
    [ME_GO_D_RefID] int NULL,
    [QtyRcvd] int NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EmpGatePass')
BEGIN
    CREATE TABLE [dbo].[EmpGatePass] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [EmpID] varchar(50) NULL,
    [DT] datetime NULL,
    [TimeOut] char(5) NULL,
    [TimeIn] char(5) NULL,
    [EmpID_Recommendation] varchar(50) NULL,
    [EmpID_Approval] varchar(50) NULL,
    [Purpose] varchar(1000) NULL,
    [GPType] tinyint NULL,
    [TimeTotal] char(5) NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL DEFAULT (getdate()),
    [EmpTimes_RefID] int NULL,
    [HP_TimeOut] char(5) NULL,
    [HP_TimeIn] char(5) NULL,
    [EmpTimes_In_RefID] int NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EmpGatePass_ApprovalAuthorities')
BEGIN
    CREATE TABLE [dbo].[EmpGatePass_ApprovalAuthorities] (
    [EmpID] varchar(50) NOT NULL,
    CONSTRAINT [PK_EmpGatePass_ApprovalAuthorities] PRIMARY KEY CLUSTERED ([EmpID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MaterialInward')
BEGIN
    CREATE TABLE [dbo].[MaterialInward] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [PoNumber] varchar(50) NULL,
    [SupplierName] varchar(150) NULL,
    [VehicleEntryId] int NULL,
    [VehicleNumber] varchar(50) NULL,
    [DriverName] varchar(100) NULL,
    [DeliveryChallanNo] varchar(50) NULL,
    [InvoiceNo] varchar(50) NULL,
    [Weighment] decimal(18,3) NULL,
    [EntryTime] datetime NOT NULL DEFAULT (getdate()),
    [Status] varchar(20) NOT NULL DEFAULT ('Open'),
    [CreatedBy] varchar(50) NULL,
    [UpdatedBy] varchar(50) NULL,
    [CreatedAt] datetime NOT NULL DEFAULT (getdate()),
    [UpdatedAt] datetime NULL,
    [Remarks] varchar(255) NULL,
    CONSTRAINT [PK_MaterialInward] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MaterialInwardLines')
BEGIN
    CREATE TABLE [dbo].[MaterialInwardLines] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [MaterialInwardId] int NOT NULL,
    [MaterialName] varchar(150) NOT NULL,
    [Quantity] decimal(18,3) NOT NULL DEFAULT ((0)),
    [Uom] varchar(25) NULL,
    CONSTRAINT [PK_MaterialInwardLines] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MaterialOutward')
BEGIN
    CREATE TABLE [dbo].[MaterialOutward] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [SalesOrderNo] varchar(50) NULL,
    [DeliveryOrderNo] varchar(50) NULL,
    [CustomerName] varchar(150) NULL,
    [VehicleEntryId] int NULL,
    [VehicleNumber] varchar(50) NULL,
    [DriverName] varchar(100) NULL,
    [Weighment] decimal(18,3) NULL,
    [EntryTime] datetime NOT NULL DEFAULT (getdate()),
    [Status] varchar(20) NOT NULL DEFAULT ('Dispatched'),
    [CreatedBy] varchar(50) NULL,
    [UpdatedBy] varchar(50) NULL,
    [CreatedAt] datetime NOT NULL DEFAULT (getdate()),
    [UpdatedAt] datetime NULL,
    [Remarks] varchar(255) NULL,
    CONSTRAINT [PK_MaterialOutward] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MaterialOutwardLines')
BEGIN
    CREATE TABLE [dbo].[MaterialOutwardLines] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [MaterialOutwardId] int NOT NULL,
    [MaterialName] varchar(150) NOT NULL,
    [Quantity] decimal(18,3) NOT NULL DEFAULT ((0)),
    [Uom] varchar(25) NULL,
    CONSTRAINT [PK_MaterialOutwardLines] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MiscEntries_GateInTemp')
BEGIN
    CREATE TABLE [dbo].[MiscEntries_GateInTemp] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [GatePassNo] varchar(50) NULL,
    [DT] datetime NULL,
    [PersonName] varchar(255) NULL,
    [VehicleNo] varchar(50) NULL,
    [Description] varchar(255) NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL DEFAULT (getdate()),
    [EmpID] varchar(50) NULL,
    [GatePassType] int NULL DEFAULT ((0)),
    [DeptID] varchar(50) NULL,
    [Approved_HOD] bit NULL DEFAULT ((0)),
    [Approved] bit NULL DEFAULT ((0)),
    [Approved_Finance] bit NULL DEFAULT ((0)),
    [GatePassType2] int NULL DEFAULT ((0)),
    [VendID] int NULL,
    [VendNameManual] varchar(500) NULL,
    [EmpNameManual] varchar(50) NULL,
    [GatePassTypeName] varchar(50) NULL,
    CONSTRAINT [PK_MiscEntries_GateInTemp] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MiscEntries_GateInTemp_Details')
BEGIN
    CREATE TABLE [dbo].[MiscEntries_GateInTemp_Details] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [MEGO_RefID] int NULL,
    [Detail] varchar(255) NULL,
    [Unit] varchar(50) NULL,
    [Rate] real NULL,
    [Qty] real NULL,
    [ReturnType] varchar(20) NULL,
    [ReturnDT] datetime NULL,
    [QtyRcvd] real NULL,
    [RcvdDT] datetime NULL,
    [RcvdBy] varchar(50) NULL,
    [RcvdRemarks] varchar(255) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'VehicleCategories')
BEGIN
    CREATE TABLE [dbo].[VehicleCategories] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [CategoryName] varchar(50) NOT NULL,
    [SortOrder] int NOT NULL DEFAULT ((0)),
    [Active] bit NOT NULL DEFAULT ((1)),
    CONSTRAINT [PK_VehicleCategories] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'VehicleEntries')
BEGIN
    CREATE TABLE [dbo].[VehicleEntries] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [VehicleNumber] varchar(50) NOT NULL,
    [VehicleType] varchar(50) NULL,
    [DriverName] varchar(100) NULL,
    [DriverCNIC] varchar(50) NULL,
    [DriverMobile] varchar(50) NULL,
    [SupplierCompany] varchar(150) NULL,
    [Transporter] varchar(150) NULL,
    [Purpose] varchar(255) NULL,
    [CategoryId] int NULL,
    [EntryTime] datetime NOT NULL DEFAULT (getdate()),
    [ExitTime] datetime NULL,
    [VehiclePhotoBase64] varchar(MAX) NULL,
    [Status] varchar(20) NOT NULL DEFAULT ('Inside'),
    [CreatedBy] varchar(50) NULL,
    [UpdatedBy] varchar(50) NULL,
    [CreatedAt] datetime NOT NULL DEFAULT (getdate()),
    [UpdatedAt] datetime NULL,
    [Remarks] varchar(255) NULL,
    CONSTRAINT [PK_VehicleEntries] PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Visitors')
BEGIN
    CREATE TABLE [dbo].[Visitors] (
    [EntryID] int IDENTITY(1,1) NOT NULL,
    [VisitorName] varchar(50) NULL,
    [CompanyName] varchar(50) NULL,
    [NICNumber] varchar(50) NULL,
    [EmpID] varchar(50) NULL,
    [DT] datetime NULL,
    [Purpose] varchar(255) NULL,
    [Remarks] varchar(255) NULL,
    [UserName] varchar(50) NULL,
    [MachineName] varchar(50) NULL,
    [DTEntry] datetime NULL DEFAULT (getdate()),
    [TimeOut] datetime NULL,
    [ConcernedPerson] varchar(50) NULL,
    [MobileNo] varchar(50) NULL,
    [CardNo] varchar(50) NULL,
    [VendID] int NULL,
    [EmpNameManual] varchar(50) NULL,
    [VendNameManual] varchar(50) NULL,
    [ConcernedPersonEmail] varchar(150) NULL,
    [ConcernedPersonWhatsApp] varchar(50) NULL,
    [NotifyByEmail] bit NOT NULL DEFAULT ((0)),
    [NotifyByWhatsApp] bit NOT NULL DEFAULT ((0)),
    [VisitCode] varchar(50) NULL,
    [QrPayload] varchar(255) NULL,
    [LastNotificationAt] datetime NULL,
    [NotificationStatus] varchar(50) NULL,
    [OutVerifiedBy] varchar(50) NULL,
    [OutVerifiedAt] datetime NULL,
    [OutVerificationInput] varchar(100) NULL
    );
END
GO

-- ------------------------------------------------------------------------------
-- 2. SCALAR FUNCTIONS
-- ------------------------------------------------------------------------------

CREATE OR ALTER FUNCTION [dbo].[GetTransactionValue] (@AccNo AS VARCHAR(50),@DTFrom DATETIME,@DTTo DATETIME)  
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

-- ------------------------------------------------------------------------------
-- 3. VIEWS DEFINITIONS (CREATE OR ALTER)
-- ------------------------------------------------------------------------------

CREATE OR ALTER VIEW [dbo].[VAccounts]
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

CREATE OR ALTER VIEW [dbo].[VVisitors]
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

CREATE OR ALTER VIEW dbo.VGateRcvdByPO
AS
SELECT PORefNo, WORef, MaterialID, SUM(ISNULL(QtyRcvd, 0) - ISNULL(QtyReturned, 0)) AS TotalGateRcvd
FROM dbo.VendRcvdDetailPO
WHERE (ISNULL(Checked, 0) = 0)
GROUP BY PORefNo, WORef, MaterialID;
GO

CREATE OR ALTER VIEW dbo.VGateRcvdByWO
AS
SELECT WORef, MaterialID, SUM(ISNULL(QtyRcvd, 0) - ISNULL(QtyReturned, 0)) AS TotalGateRcvd
FROM dbo.VendRcvdDetailPO
WHERE (ISNULL(Checked, 0) = 0)
GROUP BY WORef, MaterialID;
GO



CREATE OR ALTER VIEW [dbo].[VGatePass_Return]
AS
SELECT        dbo.GatePass_Return.EntryID, dbo.GatePass_Return.GatePassNo, dbo.GatePass_Return.DT, dbo.GatePass_Return.PersonName, dbo.GatePass_Return.VehicleNo, 
              dbo.GatePass_Return.Description, dbo.GatePass_Return.UserName, dbo.GatePass_Return.MachineName, dbo.GatePass_Return.DTEntry, 
              CAST(CONVERT(VARCHAR(50), dbo.GatePass_Return.DT, 6) AS DATETIME) AS OnlyDT, dbo.GatePass_Return.SenderName, dbo.GatePass_Return.Returnable, 
              dbo.GatePass_Return.IssdTo, dbo.GatePass_Return.GatePassType, dbo.GatePass_Return.ReturnDT, T1.TotalQty, T1.TotalQtyRcvd, 
              dbo.GatePass_Return.DeptID_From, dbo.GatePass_Return.DeptID_To, dbo.Departments.name AS DeptName_From, 
              CASE WHEN DeptID_To = '0' THEN 'Outdoor Issuance' ELSE Departments_1.name END AS DeptName_To, dbo.GatePass_Return.VendID, 
              ISNULL(vac.AccTitle, '') AS Name, 
              dbo.GatePass_Return.ApprovedBy_Name, dbo.GatePass_Return.ApprovedBy_Designation
FROM          dbo.Departments RIGHT OUTER JOIN
              dbo.GatePass_Return LEFT OUTER JOIN
              dbo.Venders ON dbo.GatePass_Return.VendID = dbo.Venders.VendID LEFT OUTER JOIN
              dbo.Accounts vac ON dbo.Venders.AccNo = vac.AccNo LEFT OUTER JOIN
              dbo.Departments AS Departments_1 ON dbo.GatePass_Return.DeptID_To = Departments_1.deptid ON 
              dbo.Departments.deptid = dbo.GatePass_Return.DeptID_From LEFT OUTER JOIN
              (SELECT MEGO_RefID, SUM(Qty) AS TotalQty, SUM(TotalQtyRcvd) AS TotalQtyRcvd
               FROM dbo.VGatePass_Return_Details
               GROUP BY MEGO_RefID) AS T1 ON dbo.GatePass_Return.EntryID = T1.MEGO_RefID;
GO


-- ------------------------------------------------------------------------------
-- 4. STORED PROCEDURES
-- ------------------------------------------------------------------------------

CREATE OR ALTER FUNCTION [dbo].[Get_FixedAssets_Adjustment] (@AssetID AS INT,@DT AS DATETIME)  
	RETURNS REAL AS
	BEGIN
		DECLARE @ReturnVal REAL
		DECLARE @Rate AS REAL

		SELECT @Rate=SUM(ISNULL(Adjustment,0)) FROM FixedAssets_Adjustment WHERE DT<=@DT AND AssetID=@AssetID GROUP BY AssetID
		SET @ReturnVal=@Rate

		RETURN @ReturnVal
	END
GO

CREATE OR ALTER PROCEDURE EmpGatePassSheet_SP @DTFrom DateTime,@DTTo DateTime
AS
	
BEGIN
	
	SELECT VEmp.EmpID,Name,Rel,FName,StartingSalary,DeptName,DeptID,Designation,Active,Tab2.TotalGPHrs FROM VEmp
		INNER JOIN (SELECT EmpID,SUM(GPHrs) AS TotalGPHrs FROM EmpGPHrs WHERE DT BETWEEN @DTFrom AND @DTTo GROUP BY EmpID) Tab2
	 ON VEmp.EmpID=Tab2.EmpID
	 ORDER BY VEmp.EmpID
END
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Financial_Budgeting_Report_Glance] (@FinancialYear VARCHAR(50),@DTFrom AS DATETIME,@DTTo AS DATETIME)
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

CREATE OR ALTER PROCEDURE [dbo].[FinancialBudgeting_Expenses_Report_SP] 
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

CREATE OR ALTER PROCEDURE [dbo].[FinancialBudgeting_Expenses_SP] (@FinancialYear VARCHAR(50),@DTFrom As DateTime,@DTTo As DateTime,@Type AS VARCHAR(50)='0') AS

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

CREATE OR ALTER PROCEDURE [dbo].[FinancialBudgeting_Monthwise_SP] (@FinancialYear VARCHAR(50),@DTFrom As DateTime,@DTTo As DateTime,@Type VARCHAR(50)='0') AS

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

CREATE OR ALTER PROCEDURE [dbo].[FinancialBudgeting_Report_SP]
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

-- Synchronize Default Permissions for Admin and Accounts
UPDATE dbo.Users SET BudgetingMainLink = 1, FixedAssetsMainLink = 1, GateMainLink = 1 WHERE UserName IN ('Administrator', 'accounts', 'Usman');
UPDATE dbo.Users SET BudgetingMainLink = COALESCE(BudgetingMainLink, 0), FixedAssetsMainLink = COALESCE(FixedAssetsMainLink, 0), GateMainLink = COALESCE(GateMainLink, 0) WHERE BudgetingMainLink IS NULL OR FixedAssetsMainLink IS NULL OR GateMainLink IS NULL;
GO

