USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VFCustomers]    Script Date: 10/5/2026 1:09:44 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VFCustomers]
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


