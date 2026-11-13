
/****** Object:  View [dbo].[VRunningLots_POI]    Script Date: 11/4/2026 8:05:28 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



ALTER VIEW [dbo].[VRunningLots_POI]
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
                             (SELECT        ISNULL(T1.StartDT, T2.StartDT) AS StartDT, dbo.RunningLots_Temp.LotNo, dbo.RunningLots_Temp.ItemCode, dbo.RunningLots_Temp.Qty, dbo.RunningLots_Temp.Description, DATEDIFF(day, 
                                                         dbo.RunningLots_Temp.DT, GETDATE()) AS LastProcessDays, dbo.RunningLots_Temp.OrderNo, '' AS Remarks, dbo.RunningLots_Temp.DT, dbo.RunningLots_Temp.ProcessID, 
                                                         dbo.RunningLots_Temp.ReWorkLot, 0 AS EntryType, dbo.POI_Remarks.Remarks AS POIRemarks, 0 AS OpenPOQty, dbo.RunningLots_Temp.VendID, 
                                                         CASE WHEN SNo <= PolishingItemSNo THEN 1 ELSE 0 END AS R1, dbo.VLotsWithMasterPONo.MasterPONo,dbo.RunningLots_Temp.TrayNo
                               FROM            dbo.RunningLots_Temp LEFT OUTER JOIN
                                                         dbo.POI_Remarks ON dbo.RunningLots_Temp.LotNo = dbo.POI_Remarks.LotNo LEFT OUTER JOIN
                                                             (SELECT        dbo.vendRcvdDetail.LotNo, MIN(dbo.VendReceived.DT) AS StartDT
                                                               FROM            dbo.VendReceived INNER JOIN
                                                                                         dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID
                                                               GROUP BY dbo.vendRcvdDetail.LotNo) AS T1 ON dbo.RunningLots_Temp.LotNo = T1.LotNo INNER JOIN
                                                         dbo.VItems ON dbo.RunningLots_Temp.ItemCode = dbo.VItems.ItemID LEFT OUTER JOIN
                                                         dbo.FCustomerOrders ON dbo.RunningLots_Temp.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                                                             (SELECT        dbo.VendIssdDetail.LotNo, MIN(dbo.VendIssued.DT) AS StartDT
                                                               FROM            dbo.VendIssued INNER JOIN
                                                                                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID
                                                               GROUP BY dbo.VendIssdDetail.LotNo) AS T2 ON dbo.RunningLots_Temp.LotNo = T2.LotNo LEFT OUTER JOIN
                                                         dbo.VLotsWithMasterPONo ON dbo.RunningLots_Temp.LotNo = dbo.VLotsWithMasterPONo.LotNo
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
                                                              FROM            dbo.RunningLots_Temp AS VRunningLots_Simple_3
                                                              GROUP BY OrderNo, ItemCode) AS T1_1 ON dbo.VUnshippedOrderItemList.OrderNo = T1_1.OrderNo AND dbo.VUnshippedOrderItemList.CompItemCode = T1_1.ItemCode) AS TUnion INNER JOIN
                         dbo.VItems AS VItems_1 ON TUnion.ItemCode = VItems_1.ItemID INNER JOIN
                         (SELECT OrderNo,CompItemCode,MAX(Price) AS Price,MAX(DeliveryDT) AS DeliveryDT,SUM(FOrderItems.Qty) AS OrderQty FROM FOrderItems GROUP BY OrderNo,CompItemCode) FOrderItems  ON TUnion.ItemCode = FOrderItems.CompItemCode AND TUnion.OrderNo = FOrderItems.OrderNo ON FCustomerOrders_1.OrderNo = TUnion.OrderNo LEFT OUTER JOIN
                             (SELECT        OrderNo, ItemCode, SUM(Qty) AS TLRunning
                               FROM            dbo.RunningLots_Temp AS VRunningLots_Simple_2
                               GROUP BY OrderNo, ItemCode) AS TRunning ON TUnion.OrderNo = TRunning.OrderNo AND TUnion.ItemCode = TRunning.ItemCode LEFT OUTER JOIN
                             (SELECT        OrderNo, ItemCode, SUM(Qty) AS TLRunning_Order
                               FROM            dbo.RunningLots_Temp AS VRunningLots_Simple_1
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
