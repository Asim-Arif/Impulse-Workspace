IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_MakerPostedBillsDetail_Receivings_VRDRefID' AND object_id = OBJECT_ID(N'dbo.MakerPostedBillsDetail_Receivings'))
    CREATE NONCLUSTERED INDEX IX_MakerPostedBillsDetail_Receivings_VRDRefID
    ON dbo.MakerPostedBillsDetail_Receivings (VRD_RefID) INCLUDE (EntryID, MPB_RefID);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_DispatchListDetail_VRD_VRDRefID' AND object_id = OBJECT_ID(N'dbo.DispatchListDetail_VRD'))
    CREATE NONCLUSTERED INDEX IX_DispatchListDetail_VRD_VRDRefID
    ON dbo.DispatchListDetail_VRD (VRD_RefID) INCLUDE (Qty, DLDC_RefID);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_LotTransferDetails_Type_VRDFromRefID' AND object_id = OBJECT_ID(N'dbo.LotTransferDetails'))
    CREATE NONCLUSTERED INDEX IX_LotTransferDetails_Type_VRDFromRefID
    ON dbo.LotTransferDetails (Type, VRD_From_RefID);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_VendIssdDetail_RcvdRefID' AND object_id = OBJECT_ID(N'dbo.VendIssdDetail'))
    CREATE NONCLUSTERED INDEX IX_VendIssdDetail_RcvdRefID
    ON dbo.VendIssdDetail (Rcvd_RefID);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_VendRcvdDetail_IssueRefID_RcvdQty' AND object_id = OBJECT_ID(N'dbo.vendRcvdDetail'))
    CREATE NONCLUSTERED INDEX IX_VendRcvdDetail_IssueRefID_RcvdQty
    ON dbo.vendRcvdDetail (Issue_RefID) INCLUDE (RcvdQty);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_TransferredToReadyFinishLotsDetail_VRDRefID_Inc' AND object_id = OBJECT_ID(N'dbo.TransferredToReadyFinishLotsDetail'))
    CREATE NONCLUSTERED INDEX IX_TransferredToReadyFinishLotsDetail_VRDRefID_Inc
    ON dbo.TransferredToReadyFinishLotsDetail (VRD_RefID) INCLUDE (Qty, RefID, EntryID);
GO

CREATE OR ALTER VIEW [dbo].[VVendReceivingList]
AS
SELECT
    dbo.VendReceived.EntryID,
    dbo.VendReceived.VendID,
    dbo.VendReceived.DT,
    dbo.VendReceived.UserID,
    dbo.VendReceived.ProcessID,
    dbo.vendRcvdDetail.RecieptID,
    dbo.vendRcvdDetail.ItemCode,
    dbo.vendRcvdDetail.RcvdQty,
    dbo.vendRcvdDetail.Wastage,
    dbo.vendRcvdDetail.LostQty,
    dbo.Makers.VenderName,
    dbo.Makers.VendID1,
    CAST(CONVERT(varchar(10), dbo.VendReceived.DT, 1) AS datetime) AS OnlyDT,
    dbo.vendRcvdDetail.ReqAuth,
    dbo.VendReceived.Issuance_RefID,
    dbo.Processes.Description,
    dbo.ItemProcesses.Scanning,
    dbo.vendRcvdDetail.LotNo,
    dbo.vendRcvdDetail.NextProcessID,
    dbo.vendRcvdDetail.IssQty,
    dbo.vendRcvdDetail.EntryID AS VRD_EntryID,
    dbo.vendRcvdDetail.Opening_RefID,
    dbo.Items.CatID,
    dbo.vendRcvdDetail.OrderNo,
    dbo.VendIssued.MasterPONo,
    dbo.vendRcvdDetail.ReWorkLot,
    FOrderItems.CompItemCode,
    dbo.vendRcvdDetail.ReWorkQty,
    dbo.FCustomerOrders.CustCode,
    dbo.Items.ItemName,
    dbo.Items.ItemSize,
    dbo.Items.SizeUnit,
    dbo.Items.Type,
    dbo.Items.GroupID,
    dbo.ItemGroups.Description AS ItemGroup,
    dbo.VendReceived.EmpID,
    T1.VRD_From_RefID,
    dbo.VendIssued.DT AS IssDT,
    dbo.VendIssdDetail.IssQty AS Issuance_IssQty,
    dbo.VendIssdDetail.RcvdQty AS Issuance_RcvdQty,
    TEmp.Empoloyees,
    TDL.VRD_RefID,
    dbo.VendIssdDetail.ReturnDT,
    dbo.MakerPostedBillsDetail_Receivings.EntryID AS MPB_D_EntryID,
    dbo.MakerPostedBills.BillNo,
    dbo.FCustomerOrders.InternalRefNo,
    FOrderItems.OrderQty,
    dbo.VendIssued.Closed,
    dbo.VItems_Complaints.ItemID AS ComplaintItemID,
    dbo.Lots_List.Batch_No,
    dbo.Lots_List.Mill_Certificate_No,
    dbo.Items.TipSize,
    dbo.VendReceived.EntryDT,
    dbo.Employees.Phone1,
    dbo.VHubSupervisors.Hub_Name,
    dbo.VHubSupervisors.Supervisors,
    ISNULL(TDL.DispatchQty, 0) AS DispatchQty,
    ISNULL(TTRF.StoreTransferQty, 0) AS StoreTransferQty,
    CASE WHEN TLC.LotNo IS NOT NULL THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS IsLotClosed,
    dbo.VendIssued.BillingDays,
    dbo.VendIssdDetail.Rate,
    TLotDispatch.DispatchListNo,
    TTRFLD.TTRFD_EntryID,
    TRunLot.Lot_Current_Process,
    TTotalRcvd.Total_Rcvd_Qty,
    TTRFLD.Shelf_RefID,
    TTRFLD.StoreName,
    TTRFLD.RackNo,
    TTRFLD.ShelfNo
FROM dbo.VendReceived
INNER JOIN dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID
INNER JOIN dbo.Makers ON dbo.VendReceived.VendID = dbo.Makers.VendID
INNER JOIN dbo.Processes ON dbo.VendReceived.ProcessID = dbo.Processes.ProcessID
INNER JOIN dbo.Items ON dbo.vendRcvdDetail.ItemCode = dbo.Items.ItemID
LEFT OUTER JOIN dbo.Employees ON dbo.VendReceived.EmpID = dbo.Employees.empid
LEFT OUTER JOIN dbo.Lots_List ON dbo.vendRcvdDetail.LotNo = dbo.Lots_List.LotNo
LEFT OUTER JOIN dbo.FCustomerOrders ON dbo.vendRcvdDetail.OrderNo = dbo.FCustomerOrders.OrderNo
LEFT OUTER JOIN dbo.VItems_Complaints ON dbo.Items.ItemID = dbo.VItems_Complaints.ItemID
LEFT OUTER JOIN dbo.MakerPostedBillsDetail_Receivings ON dbo.vendRcvdDetail.EntryID = dbo.MakerPostedBillsDetail_Receivings.VRD_RefID
LEFT OUTER JOIN dbo.MakerPostedBills ON dbo.MakerPostedBillsDetail_Receivings.MPB_RefID = dbo.MakerPostedBills.EntryID
LEFT OUTER JOIN dbo.VendIssdDetail ON dbo.vendRcvdDetail.Issue_RefID = dbo.VendIssdDetail.EntryID
LEFT OUTER JOIN dbo.VendIssued ON dbo.VendIssdDetail.RefID = dbo.VendIssued.EntryID
LEFT OUTER JOIN dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID
LEFT OUTER JOIN dbo.ItemProcesses
    ON dbo.Processes.ProcessID = dbo.ItemProcesses.ProcessID
   AND dbo.vendRcvdDetail.ItemCode = dbo.ItemProcesses.ItemID
   AND dbo.VendReceived.ProcessID = dbo.ItemProcesses.ProcessID
LEFT OUTER JOIN dbo.Lots_Closed AS TLC ON dbo.vendRcvdDetail.LotNo = TLC.LotNo
LEFT OUTER JOIN dbo.ItemProcessGroups ON dbo.vendRcvdDetail.ItemCode = dbo.ItemProcessGroups.ItemID
LEFT OUTER JOIN dbo.ProcessGroupsProcesses
    ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroupsProcesses.Group_RefID
   AND dbo.VendReceived.ProcessID = dbo.ProcessGroupsProcesses.Process_RefID
LEFT OUTER JOIN dbo.VHubSupervisors
    ON dbo.ProcessGroupsProcesses.Group_RefID = dbo.VHubSupervisors.GroupID
   AND dbo.ProcessGroupsProcesses.Hub_Name = dbo.VHubSupervisors.Hub_Name
OUTER APPLY (
    SELECT ISNULL(STUFF((
        SELECT ', ' + E.Name
        FROM dbo.VendReceived_Employees AS VRE
        INNER JOIN dbo.Employees AS E ON VRE.EmpID = E.empid
        WHERE VRE.VR_RefID = dbo.VendReceived.EntryID
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, ''), '') AS Empoloyees
) AS TEmp
OUTER APPLY (
    SELECT foi.CompItemCode, SUM(foi.Qty) AS OrderQty
    FROM dbo.VFOrderItems AS foi
    WHERE foi.OrderNo = dbo.vendRcvdDetail.OrderNo
      AND foi.CompItemCode = dbo.vendRcvdDetail.ItemCode
    GROUP BY foi.CompItemCode
) AS FOrderItems
OUTER APPLY (
    SELECT TOP (1) ltd.VRD_From_RefID
    FROM dbo.LotTransferDetails AS ltd
    WHERE ltd.Type = 1
      AND ltd.VRD_From_RefID = dbo.vendRcvdDetail.EntryID
) AS T1
OUTER APPLY (
    SELECT dld.VRD_RefID, SUM(dld.Qty) AS DispatchQty
    FROM dbo.DispatchListDetail_VRD AS dld
    WHERE dld.VRD_RefID = dbo.vendRcvdDetail.EntryID
    GROUP BY dld.VRD_RefID
) AS TDL
OUTER APPLY (
    SELECT MAX(dl.DispatchListNo) AS DispatchListNo
    FROM dbo.vendRcvdDetail AS vrd2
    INNER JOIN dbo.DispatchListDetail_VRD AS dld2 ON vrd2.EntryID = dld2.VRD_RefID
    INNER JOIN dbo.DispatchListDetail_Inners AS dldi ON dld2.DLDC_RefID = dldi.EntryID
    INNER JOIN dbo.DispatchListDetails_Adv AS dlda ON dldi.RefID = dlda.EntryID
    INNER JOIN dbo.DispatchList AS dl ON dlda.RefID = dl.EntryID
    WHERE vrd2.LotNo = dbo.vendRcvdDetail.LotNo
) AS TLotDispatch
OUTER APPLY (
    SELECT
        MAX(trf.EntryID) AS TTRFD_EntryID,
        MAX(pl.Shelf_RefID) AS Shelf_RefID,
        MAX(pl.StoreName) AS StoreName,
        MAX(pl.ShelfNo) AS ShelfNo,
        MAX(pl.RackNo) AS RackNo
    FROM dbo.vendRcvdDetail AS vrd3
    INNER JOIN dbo.TransferredToReadyFinishLotsDetail AS trf ON vrd3.EntryID = trf.VRD_RefID
    INNER JOIN dbo.TransferredToReadyFinishLots AS trfh ON trf.RefID = trfh.EntryID
    LEFT OUTER JOIN dbo.VRcvItemsSimpleDetail_Placement AS pl ON trfh.EntryID = pl.TTRFLD_RefID
    WHERE vrd3.LotNo = dbo.vendRcvdDetail.LotNo
) AS TTRFLD
OUTER APPLY (
    SELECT SUM(trf2.Qty) AS StoreTransferQty
    FROM dbo.TransferredToReadyFinishLotsDetail AS trf2
    WHERE trf2.VRD_RefID = dbo.vendRcvdDetail.EntryID
) AS TTRF
OUTER APPLY (
    SELECT SUM(vrd4.RcvdQty) AS Total_Rcvd_Qty
    FROM dbo.vendRcvdDetail AS vrd4
    WHERE dbo.vendRcvdDetail.Issue_RefID <> 0
      AND vrd4.Issue_RefID = dbo.vendRcvdDetail.Issue_RefID
) AS TTotalRcvd
OUTER APPLY (
    SELECT TOP (1) rl.Description AS Lot_Current_Process
    FROM dbo.VRunningLots_Simple_Main AS rl
    WHERE TLC.LotNo IS NULL
      AND dbo.vendRcvdDetail.LotNo <> '0'
      AND rl.LotNo = dbo.vendRcvdDetail.LotNo
    ORDER BY rl.SNO DESC
) AS TRunLot
GO