USE [SMBI_AWM]
GO

/****** Object:  Table [dbo].[ForeignCustomers]    Script Date: 6/22/2026 7:46:36 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ForeignCustomers](
	[CustCode] [varchar](50) NOT NULL,
	[Country] [varchar](50) NOT NULL,
	[Name] [varchar](1000) NULL,
	[Address] [varchar](1000) NULL,
	[FakeAddress] [varchar](1000) NULL,
	[Phone1] [varchar](50) NULL,
	[Phone2] [varchar](50) NULL,
	[Phone3] [varchar](50) NULL,
	[Email1] [varchar](100) NULL,
	[Email2] [varchar](100) NULL,
	[Fax1] [varchar](50) NULL,
	[Fax2] [varchar](50) NULL,
	[Fax3] [varchar](50) NULL,
	[URL] [varchar](100) NULL,
	[Curr] [varchar](50) NULL,
	[Cont1name] [varchar](50) NULL,
	[cont1phone] [varchar](50) NULL,
	[cont1email] [varchar](100) NULL,
	[cont1Mobile] [varchar](50) NULL,
	[Cont1Skype] [varchar](50) NULL,
	[cont2name] [varchar](50) NULL,
	[cont2phone] [varchar](50) NULL,
	[cont2email] [varchar](100) NULL,
	[cont2Mobile] [varchar](50) NULL,
	[Cont2Skype] [varchar](50) NULL,
	[AgentName] [varchar](50) NULL,
	[AgentAddress] [varchar](255) NULL,
	[AgentPhone1] [varchar](50) NULL,
	[AgentPhone2] [varchar](50) NULL,
	[AgentFax1] [varchar](50) NULL,
	[AgentFax2] [varchar](50) NULL,
	[AgentEmail] [varchar](50) NULL,
	[AgentWeb] [varchar](100) NULL,
	[ACName] [varchar](50) NULL,
	[ACPhone] [varchar](50) NULL,
	[ACMobile] [varchar](50) NULL,
	[AgentURL] [varchar](50) NULL,
	[ACEmail] [varchar](50) NULL,
	[DTFormat] [varchar](50) NULL,
	[Active] [bit] NOT NULL,
	[AccNo] [varchar](255) NULL,
	[NearestAirport] [varchar](50) NULL,
	[NearestRailwaystation] [varchar](50) NULL,
	[FedexNo] [varchar](50) NULL,
	[DHLNo] [varchar](50) NULL,
	[TradeTerms] [varchar](50) NULL,
	[PaymentTerms] [varchar](2000) NULL,
	[AcceptsExtraQty] [bit] NULL,
	[SpecialInstructions] [varchar](4000) NULL,
	[City] [varchar](255) NULL,
	[MaxDiscount] [real] NULL,
	[PaymentDays] [int] NULL,
	[LateOrderAlerts] [bit] NULL,
	[Customer_Type] [varchar](50) NULL,
	[OuterPackingLabel] [varchar](255) NULL,
	[InnerPackingLabel] [varchar](255) NULL,
	[ShowCustomerRef] [bit] NULL,
	[FinishingQuality] [varchar](50) NULL,
	[Stamps] [varchar](50) NULL,
	[RcvdVia] [varchar](50) NULL,
	[DefaultPort] [int] NULL,
	[DefaultShipMethod] [varchar](50) NULL,
	[Cont1Designation] [varchar](50) NULL,
	[Cont2Designation] [varchar](50) NULL,
	[CustomerSource] [varchar](50) NULL,
	[FC_Note_I] [varchar](4000) NULL,
	[FC_Note_II] [varchar](4000) NULL,
	[Inner_Label_Manual_I] [varchar](255) NULL,
	[Inner_Label_Manual_II] [varchar](255) NULL,
 CONSTRAINT [PK_Customers] PRIMARY KEY CLUSTERED 
(
	[CustCode] ASC,
	[Country] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ForeignCustomers] ADD  CONSTRAINT [DF_ForeignCustomers_Active]  DEFAULT (0) FOR [Active]
GO

ALTER TABLE [dbo].[ForeignCustomers] ADD  CONSTRAINT [DF_ForeignCustomers_AcceptsExtraQty]  DEFAULT (0) FOR [AcceptsExtraQty]
GO

ALTER TABLE [dbo].[ForeignCustomers] ADD  CONSTRAINT [DF_ForeignCustomers_LateOrderAlerts]  DEFAULT (0) FOR [LateOrderAlerts]
GO

ALTER TABLE [dbo].[ForeignCustomers] ADD  CONSTRAINT [DF_ForeignCustomers_ShowCustomerRef]  DEFAULT (0) FOR [ShowCustomerRef]
GO

