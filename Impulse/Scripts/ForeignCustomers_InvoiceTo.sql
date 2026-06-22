USE [SMBI_AWM]
GO

/****** Object:  Table [dbo].[ForeignCustomers_InvoiceTo]    Script Date: 6/22/2026 7:47:46 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ForeignCustomers_InvoiceTo](
	[EntryID] [int] IDENTITY(1,1) NOT NULL,
	[CustCode] [varchar](50) NULL,
	[Country] [varchar](50) NULL,
	[InvoiceTo] [varchar](1000) NULL
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ForeignCustomers_InvoiceTo]  WITH NOCHECK ADD  CONSTRAINT [FK_ForeignCustomers_InvoiceTo_ForeignCustomers] FOREIGN KEY([CustCode], [Country])
REFERENCES [dbo].[ForeignCustomers] ([CustCode], [Country])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[ForeignCustomers_InvoiceTo] CHECK CONSTRAINT [FK_ForeignCustomers_InvoiceTo_ForeignCustomers]
GO

