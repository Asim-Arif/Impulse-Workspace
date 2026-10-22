
GO
/****** Object:  StoredProcedure [dbo].[Track_Remaining_Qty_SP]    Script Date: 10/22/2026 6:06:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Track_Remaining_Qty_SP] (@OrderNo VARCHAR(50))
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

SELECT FCO.CustCode,FCO.Country,FCO.OrderNO,FOI.CompItemCode,FOI.ItemCode,VItems.ItemName,FOI.Qty AS OrderQty
	,TSOA.SOA_Qty,TRcv.Received_Qty,TShipped.Shipped_Qty,FOI.Qty-ISNULL(TShipped.Shipped_Qty,0) AS Remaining_Qty
	,TMTS.Move_To_Stock_Qty
	,TRR.Rework_Qty,TRR.Rejection_Qty
	,FCO.DT AS OrderDT
	,VItems.InHand
	,TShipped.Shipped_Qty AS ItemShippedQty
	FROM FCustomerOrders FCO
	INNER JOIN FOrderItems FOI ON FCO.OrderNo=FOI.OrderNo
	INNER JOIN VItems ON FOI.CompItemCode=VItems.ItemID

	LEFT JOIN (SELECT OrderNo,ItemID,SUM(Qty) AS SOA_Qty FROM StockOrderAdjustments WHERE OrderNo=@OrderNo GROUP BY OrderNo,ItemID) TSOA ON FOI.OrderNo=TSOA.OrderNo AND FOI.CompItemCode=TSOA.ItemID
	LEFT JOIN (SELECT OrderNo,ItemCode,SUM(RcvdQty) AS Received_Qty FROM VendRcvdDetail 
					INNER JOIN VendReceived ON VendReceived.EntryID=VendRcvdDetail.RefID 
					INNER JOIN VendIssued ON VendReceived.Issuance_RefID=VendIssued.EntryID
					WHERE OrderNo=@OrderNo AND MasterPONo IS NULL GROUP BY OrderNo,ItemCode) TRcv ON FOI.OrderNo=TRcv.OrderNo AND FOI.CompItemCode=TRcv.ItemCode

	LEFT JOIN (SELECT OrderNo,ItemID,SUM(Qty) AS Shipped_Qty FROM VCustomInvoiceItems WHERE OrderNo=@OrderNO GROUP BY OrderNo,ItemID) TShipped ON FOI.OrderNo=TShipped.OrderNo AND FOI.CompItemCode=TShipped.ItemID
	LEFT JOIN (SELECT OrderNo,ItemCode,SUM(TTRFLD.Qty) AS Move_To_Stock_Qty FROM TransferredToReadyFinishLotsDetail TTRFLD
					INNER JOIN vendRcvdDetail VRD ON TTRFLD.VRD_RefID=VRD.EntryID WHERE OrderNo=@OrderNo GROUP BY OrderNo,ItemCode) TMTS ON FOI.OrderNo=TMTS.OrderNo AND FOI.CompItemCode=TMTS.ItemCode
	LEFT JOIN (SELECT OrderNo,ItemCode,SUM(ReWorkQty) AS Rework_Qty,SUM(Wastage) AS Rejection_Qty FROM VendRcvdDetail WHERE OrderNo=@OrderNo GROUP BY OrderNo,ItemCode) TRR ON FOI.OrderNo=TRR.OrderNo AND FOI.CompItemCode=TRR.ItemCode
	WHERE FCO.OrderNo=@OrderNo


    
END
