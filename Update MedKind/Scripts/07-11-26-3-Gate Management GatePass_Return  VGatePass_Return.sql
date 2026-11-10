-- ============================================================================
-- Gate Management: GatePass_Return & VGatePass_Return (Returnable Gate Pass)
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.GatePass_Return') AND name = 'VM_AccNo')
BEGIN
    ALTER TABLE dbo.GatePass_Return ADD VM_AccNo varchar(50) NULL;
END
GO

CREATE OR ALTER VIEW [dbo].[VGatePass_Return]
AS
SELECT        dbo.GatePass_Return.EntryID, dbo.GatePass_Return.GatePassNo, dbo.GatePass_Return.DT, dbo.GatePass_Return.PersonName, dbo.GatePass_Return.VehicleNo, 
              dbo.GatePass_Return.Description, dbo.GatePass_Return.UserName, dbo.GatePass_Return.MachineName, dbo.GatePass_Return.DTEntry, 
              CAST(CONVERT(VARCHAR(50), dbo.GatePass_Return.DT, 6) AS DATETIME) AS OnlyDT, dbo.GatePass_Return.SenderName, dbo.GatePass_Return.Returnable, 
              dbo.GatePass_Return.IssdTo, dbo.GatePass_Return.GatePassType, dbo.GatePass_Return.ReturnDT, T1.TotalQty, T1.TotalQtyRcvd, 
              dbo.GatePass_Return.DeptID_From, dbo.GatePass_Return.DeptID_To, dbo.Departments.name AS DeptName_From, 
              CASE WHEN DeptID_To = '0' THEN 'Outdoor Issuance' ELSE Departments_1.name END AS DeptName_To, dbo.GatePass_Return.VendID, 
              dbo.VVendersAndMakers.AccTitle AS Name, 
              dbo.GatePass_Return.ApprovedBy_Name, dbo.GatePass_Return.ApprovedBy_Designation,
              dbo.GatePass_Return.VM_AccNo
FROM          dbo.Departments RIGHT OUTER JOIN
              dbo.GatePass_Return LEFT OUTER JOIN
              dbo.VVendersAndMakers ON dbo.GatePass_Return.VM_AccNo = dbo.VVendersAndMakers.AccNo LEFT OUTER JOIN
              dbo.Departments AS Departments_1 ON dbo.GatePass_Return.DeptID_To = Departments_1.deptid ON 
              dbo.Departments.deptid = dbo.GatePass_Return.DeptID_From LEFT OUTER JOIN
              (SELECT MEGO_RefID, SUM(Qty) AS TotalQty, SUM(TotalQtyRcvd) AS TotalQtyRcvd
               FROM dbo.VGatePass_Return_Details
               GROUP BY MEGO_RefID) AS T1 ON dbo.GatePass_Return.EntryID = T1.MEGO_RefID;
GO