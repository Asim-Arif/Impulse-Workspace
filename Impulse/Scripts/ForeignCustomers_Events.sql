USE [SMBI_AWM]
GO

/****** Object:  Table [dbo].[ForeignCustomers_Events]    Script Date: 6/22/2026 7:47:34 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ForeignCustomers_Events](
	[EntryID] [int] IDENTITY(1,1) NOT NULL,
	[CustCode] [varchar](50) NULL,
	[Country] [varchar](50) NULL,
	[EventDT] [datetime] NULL,
	[RepeatedType] [tinyint] NULL,
	[Description] [varchar](255) NULL,
 CONSTRAINT [IX_ForeignCustomers_Events] UNIQUE NONCLUSTERED 
(
	[EntryID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ForeignCustomers_Events] ADD  CONSTRAINT [DF_ForeignCustomers_Events_Repeated]  DEFAULT (0) FOR [RepeatedType]
GO

ALTER TABLE [dbo].[ForeignCustomers_Events]  WITH CHECK ADD  CONSTRAINT [FK_ForeignCustomers_Events_ForeignCustomers] FOREIGN KEY([CustCode], [Country])
REFERENCES [dbo].[ForeignCustomers] ([CustCode], [Country])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[ForeignCustomers_Events] CHECK CONSTRAINT [FK_ForeignCustomers_Events_ForeignCustomers]
GO

