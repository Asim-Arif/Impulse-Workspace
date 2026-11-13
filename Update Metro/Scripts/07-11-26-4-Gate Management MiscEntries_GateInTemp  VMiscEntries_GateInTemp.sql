
-- ============================================================================
-- Gate Management: MiscEntries_GateInTemp & VMiscEntries_GateInTemp (Gate In)
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateInTemp') AND name = 'VM_AccNo')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateInTemp ADD VM_AccNo varchar(50) NULL;
END
GO

CREATE OR ALTER VIEW [dbo].[VMiscEntries_GateInTemp]
AS
SELECT 
    m.EntryID, 
    m.GatePassNo, 
    m.DT, 
    m.PersonName, 
    m.VehicleNo, 
    m.Description, 
    m.UserName, 
    m.MachineName, 
    m.DTEntry, 
    CAST(CONVERT(VARCHAR(50), m.DT, 6) AS DATETIME) AS OnlyDT, 
    m.Returnable, 
    m.GatePassType, 
    m.DeptID_To, 
    d.name AS DeptName_To, 
    m.EmpID, 
    emp.Name AS EmpName,
    m.VendID, 
    vm.AccTitle AS Name, 
    m.VendNameManual,
    m.VM_AccNo,
    ISNULL(m.Approved_HOD, 0) AS Approved_HOD,
    ISNULL(m.Approved, 0) AS Approved,
    ISNULL(m.Approved_Finance, 0) AS Approved_Finance
FROM dbo.MiscEntries_GateInTemp m
LEFT OUTER JOIN dbo.Departments d ON m.DeptID_To = d.deptid
LEFT OUTER JOIN dbo.Employees emp ON m.EmpID = emp.EmpID
LEFT OUTER JOIN dbo.VVendersAndMakers vm ON m.VM_AccNo = vm.AccNo;
GO
