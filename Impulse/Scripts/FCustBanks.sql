USE [SMBI_AWM]
GO

/****** Object:  Table [dbo].[FCustBanks]    Script Date: 6/22/2026 7:46:54 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[FCustBanks](
	[CustCode] [varchar](50) NOT NULL,
	[Country] [varchar](50) NULL,
	[BankID] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[BankName] [varchar](100) NULL,
	[BankCountry] [varchar](50) NULL,
	[Phone1] [varchar](50) NULL,
	[Phone2] [varchar](50) NULL,
	[Fax1] [varchar](50) NULL,
	[Fax2] [varchar](50) NULL,
	[Email1] [varchar](50) NULL,
	[Email2] [varchar](50) NULL,
	[Address1] [varchar](255) NULL,
	[Address2] [varchar](255) NULL,
	[DefaultBank] [bit] NOT NULL,
	[EntryID] [int] NULL
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[FCustBanks] ADD  CONSTRAINT [DF_FCustBanks_DefaultBank]  DEFAULT (0) FOR [DefaultBank]
GO

