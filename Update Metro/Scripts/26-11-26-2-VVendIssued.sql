USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VVendIssued]    Script Date: 11/26/2026 4:48:09 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO






ALTER     VIEW [dbo].[VVendIssued]
AS
SELECT     dbo.VendIssued.VendID, dbo.VendIssued.DT, dbo.VendIssued.RecieptID, dbo.VendIssued.ItemID, dbo.Makers.VendID1, dbo.Makers.VenderName, CAST(CONVERT(Varchar(10), dbo.VendIssued.DT, 1) AS DATETIME) AS OnlyDT, TabValue.TotalValue, 
                  dbo.VendIssued.Authorized, dbo.MakerPOPromises_F(dbo.VendIssued.EntryID) AS MakerPOReturnDTs, dbo.Processes.Description, dbo.Items.ItemName, TabValue.TotalIssQty, TabValue.LotNo, dbo.VendIssued.MasterPONo, TabValue.RcvdQty, dbo.Items.CatID, 
                  TabValue.ReWorkLot, dbo.VendIssued.ProcessID, dbo.Items.GroupID, TabValue.OrderNo, dbo.VendIssued.Closed, dbo.FCustomerOrders.CustCode, TabValue.LastRcvDT, TabValue.EntryID AS DetailEntryID, dbo.FCustomerOrders.InternalRefNo, 
                  TabValue.BookMarkEntryID, dbo.VendIssued.EntryID, dbo.Items.TipSize, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Makers.Phone1, dbo.Makers.Phone2, dbo.Makers.CPhone, dbo.Makers.ContactPerson, dbo.Makers.Address, dbo.Makers.Mobile, 
                  dbo.Makers.CompanyName, dbo.Makers.Maker_Second_Name, dbo.Makers.MakerNameUrdu, dbo.Makers.Fax1,
				  dbo.VendIssued.IssEmpID,dbo.VEmp.Name AS Emp_Name,VendIssued.DTEntry
                  ,TabValue.Batch_No
                  ,VHubSupervisors.Hub_Name,VHubSupervisors.Supervisors
                  ,TabValue.TotalOrderQty
FROM        dbo.VendIssued INNER JOIN
                  dbo.Items ON dbo.VendIssued.ItemID = dbo.Items.ItemID INNER JOIN
                  dbo.Processes ON dbo.VendIssued.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
                  dbo.Makers ON dbo.VendIssued.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                  dbo.FCustomerOrders RIGHT OUTER JOIN
                      (SELECT     dbo.VVendIssdDetailWithValueWORcving.OrderNo, dbo.VVendIssdDetailWithValueWORcving.EntryID, dbo.VVendIssdDetailWithValueWORcving.RefID, dbo.VVendIssdDetailWithValueWORcving.ReWorkLot, 
                                         dbo.VVendIssdDetailWithValueWORcving.BookMarkEntryID, SUM(dbo.VVendIssdDetailWithValueWORcving.IssValue) AS TotalValue, SUM(dbo.VVendIssdDetailWithValueWORcving.IssQty) AS TotalIssQty, 
                                         MAX(dbo.VVendIssdDetailWithValueWORcving.LotNo) AS LotNo, SUM(TRcv.RcvRcvdQty) AS RcvdQty, MAX(TRcv.RcvDT) AS LastRcvDT
                                         ,dbo.VVendIssdDetailWithValueWORcving.Batch_No
                                         ,SUM(OrderQty) AS TotalOrderQty
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


