
CREATE OR ALTER VIEW [dbo].[VRawMaterialIssuanceList]
AS
SELECT        dbo.RawMaterialIssuance.IssNo, dbo.RawMaterialIssuance.DT, dbo.RawMaterialIssuance.EntryDT, dbo.RawMaterialIssuance.UserName, dbo.RawMaterialIssuance.ComputerName, dbo.RawMaterialIssuance.VI_RefID, 
                         dbo.RawMaterialIssuance.IssdToEmpID, dbo.RawMaterialIssuance.VchrNo, dbo.RawMaterialIssuance.VO_OrderNo, dbo.RawMaterialIssuance.AuthorizedByEmpID, dbo.RawMaterialIssuance.EntryID, 
                         dbo.VEmp.name AS IssdToEmpName, dbo.Makers.VenderName, TInner1.RM, dbo.VEmp.deptid, dbo.VendIssued.VendID, dbo.RawMaterialIssuance.Remarks, dbo.RawMaterialIssuance.Entry_Type, 
                         dbo.RawMaterialIssuance.Vendor_AccNo, dbo.RawMaterialIssuance.Vendor_PO_No, dbo.RawMaterialIssuance.Vendor_VchrNo,
                         VVendersAndMakers.AccTitle AS Debit_Note_Vendor,
                         TInner1.Total_Issued_Qty,
                         TEmpID.IssdToEmpIDOnly
FROM            dbo.Makers RIGHT OUTER JOIN
                         dbo.VendIssued ON dbo.Makers.VendID = dbo.VendIssued.VendID RIGHT OUTER JOIN
                         dbo.RawMaterialIssuance ON dbo.VendIssued.EntryID = dbo.RawMaterialIssuance.VI_RefID
                         OUTER APPLY (SELECT CASE WHEN CHARINDEX('}', dbo.RawMaterialIssuance.IssdToEmpID) > 2 
                                                  THEN SUBSTRING(dbo.RawMaterialIssuance.IssdToEmpID, 2, CHARINDEX('}', dbo.RawMaterialIssuance.IssdToEmpID) - 2)
                                                  ELSE dbo.RawMaterialIssuance.IssdToEmpID END AS IssdToEmpIDOnly) TEmpID
                         LEFT OUTER JOIN
                         dbo.VEmp ON TEmpID.IssdToEmpIDOnly = dbo.VEmp.empid LEFT OUTER JOIN
                             (SELECT        IssNo, MAX('{' + RMID1 + '} ' + RMName) AS RM, SUM(Qty) AS Total_Issued_Qty
                               FROM            dbo.VRawMaterialIssuance
                               GROUP BY IssNo) AS TInner1 ON dbo.RawMaterialIssuance.IssNo = TInner1.IssNo
LEFT JOIN VVendersAndMakers ON dbo.RawMaterialIssuance.Vendor_AccNo=VVendersAndMakers.AccNo