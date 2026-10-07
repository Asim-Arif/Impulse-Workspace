USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VVendReceivingList]    Script Date: 10/7/2026 7:00:34 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER   VIEW [dbo].[VVendReceivingList]
AS
SELECT     dbo.VendReceived.EntryID, dbo.VendReceived.VendID, dbo.VendReceived.DT, dbo.VendReceived.UserID, dbo.VendReceived.ProcessID, dbo.vendRcvdDetail.RecieptID, dbo.vendRcvdDetail.ItemCode, dbo.vendRcvdDetail.RcvdQty, dbo.vendRcvdDetail.Wastage, 
                  dbo.vendRcvdDetail.LostQty, dbo.Makers.VenderName, dbo.Makers.VendID1, CAST(CAST(dbo.VendReceived.DT AS DATE) AS DATETIME) AS OnlyDT, dbo.vendRcvdDetail.ReqAuth, dbo.VendReceived.Issuance_RefID, dbo.Processes.Description, 
                  dbo.ItemProcesses.Scanning, dbo.vendRcvdDetail.LotNo, dbo.vendRcvdDetail.NextProcessID, dbo.vendRcvdDetail.IssQty, dbo.vendRcvdDetail.EntryID AS VRD_EntryID, dbo.vendRcvdDetail.Opening_RefID, dbo.Items.CatID, dbo.vendRcvdDetail.OrderNo, 
                  dbo.VendIssued.MasterPONo, dbo.vendRcvdDetail.ReWorkLot, FOrderItems.CompItemCode, dbo.vendRcvdDetail.ReWorkQty, dbo.FCustomerOrders.CustCode, dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.Items.GroupID, 
                  dbo.ItemGroups.Description AS ItemGroup, dbo.VendReceived.EmpID, T1.VRD_From_RefID, dbo.VendIssued.DT AS IssDT, dbo.VendIssdDetail.IssQty AS Issuance_IssQty, dbo.VendIssdDetail.RcvdQty AS Issuance_RcvdQty, 
                  dbo.VendReceived_Employees_F(dbo.VendReceived.EntryID) AS Empoloyees, TDL.VRD_RefID, dbo.VendIssdDetail.ReturnDT, dbo.MakerPostedBillsDetail_Receivings.EntryID AS MPB_D_EntryID, dbo.MakerPostedBills.BillNo, dbo.FCustomerOrders.InternalRefNo, 
                  FOrderItems.OrderQty, dbo.VendIssued.Closed, dbo.VItems_Complaints.ItemID AS ComplaintItemID, dbo.Lots_List.Batch_No, dbo.Lots_List.Mill_Certificate_No
                  , dbo.Items.TipSize, dbo.VendReceived.EntryDT, dbo.Employees.Phone1, 
                  VHubSupervisors.Hub_Name, VHubSupervisors.Supervisors
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
                      (SELECT     VRD_RefID
                       FROM        dbo.DispatchListDetail_VRD
                       GROUP BY VRD_RefID) AS TDL ON dbo.vendRcvdDetail.EntryID = TDL.VRD_RefID LEFT OUTER JOIN
                  dbo.VHubSupervisors ON dbo.VendReceived.EmpID = dbo.VHubSupervisors.GroupID
                  LEFT OUTER JOIN
             (SELECT LotNo, MAX(DispatchListNo) AS DispatchListNo
            FROM  dbo.VDispatchListDetail_Adv_Lot
            GROUP BY LotNo) AS TLotDispatch ON dbo.vendRcvdDetail.LotNo = TLotDispatch.LotNo LEFT OUTER JOIN
             (SELECT LotNo, MAX(EntryID) AS TTRFD_EntryID,MAX(Shelf_RefID) AS Shelf_RefID,MAX(StoreName) AS StoreName,MAX(ShelfNo) AS ShelfNo,MAX(RackNo) AS RackNo
            FROM  dbo.VTransferredToReadyFinishLotsDetail
            GROUP BY LotNo) AS TTRFLD ON dbo.vendRcvdDetail.LotNo = TTRFLD.LotNo
			LEFT JOIN VRunningLots_Simple ON dbo.vendRcvdDetail.LotNo=VRunningLots_Simple.LotNo
			LEFT JOIN (SELECT VendRcvdDetail.Issue_RefID,SUM(VendRcvdDetail.RcvdQty) AS Total_Rcvd_Qty FROM VendRcvdDetail WHERE VendRcvdDetail.Issue_RefID<>0 GROUP BY VendRcvdDetail.Issue_RefID) TTotalRcvd ON VendRcvdDetail.Issue_RefID=TTotalRcvd.Issue_RefID
GO


