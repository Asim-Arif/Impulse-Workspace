USE [SMBI_AWM]
GO

/****** Object:  Table [dbo].[ForeignCustomersShippingAddresses]    Script Date: 6/22/2026 7:47:21 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ForeignCustomersShippingAddresses](
	[EntryID] [int] IDENTITY(1,1) NOT NULL,
	[CustCode] [varchar](50) NULL,
	[Country] [varchar](50) NULL,
	[ShippingAddress] [varchar](1000) NULL
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ForeignCustomersShippingAddresses]  WITH NOCHECK ADD  CONSTRAINT [FK_ForeignCustomersShippingAddresses_ForeignCustomers] FOREIGN KEY([CustCode], [Country])
REFERENCES [dbo].[ForeignCustomers] ([CustCode], [Country])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[ForeignCustomersShippingAddresses] CHECK CONSTRAINT [FK_ForeignCustomersShippingAddresses_ForeignCustomers]
GO

