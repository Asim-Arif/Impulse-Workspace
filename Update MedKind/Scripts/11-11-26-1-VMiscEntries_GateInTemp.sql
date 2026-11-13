USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VMiscEntries_GateInTemp]    Script Date: 11/11/2026 9:56:56 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

ALTER   VIEW [dbo].[VMiscEntries_GateInTemp] AS 
SELECT m.EntryID, m.GatePassNo, m.DT, m.PersonName, m.VehicleNo, m.Description, m.UserName, m.MachineName, 
m.DTEntry, CAST(CONVERT(VARCHAR(50), m.DT, 6) AS DATETIME) AS OnlyDT, m.EmpID, emp.Name AS EmpName, m.GatePassType, m.GatePassType2, 
m.DeptID, d.name AS DeptName, m.VendID, vm.AccTitle AS Name, m.VendNameManual, m.VM_AccNo, ISNULL(m.Approved_HOD, 0) AS Approved_HOD
, ISNULL(m.Approved, 0) AS Approved, ISNULL(m.Approved_Finance, 0) AS Approved_Finance FROM dbo.MiscEntries_GateInTemp m 
LEFT OUTER JOIN dbo.Departments d ON m.DeptID = d.deptid 
LEFT OUTER JOIN dbo.Employees emp ON m.EmpID = emp.EmpID 
LEFT OUTER JOIN dbo.VVendersAndMakers vm ON m.VM_AccNo = vm.AccNo;
GO


