USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VEmpTimes2]    Script Date: 10/5/2026 1:41:08 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VEmpTimes2]
AS
SELECT     TOP 100 PERCENT dbo.Employees.empid, dbo.Employees.deptid, dbo.Employees.name, dbo.Employees.fname, dbo.Employees.Rel, dbo.Employees.Active, 
                      dbo.VEmpDailyEntries.DT, EmpTimes_1.InTime AS FirstInTime, EmpTimes_1.OutTime AS FirstOutTime, EmpTimes_2.InTime AS SecondInTime, 
                      EmpTimes_2.OutTime AS SecondOutTime, EmpTimes_3.InTime AS OTInTime, EmpTimes_3.OutTime AS OTOutTime, dbo.VEmpDailyEntries.FirstEntryID, 
                      ISNULL(EmpTimes_1.Hrs, 0) + ISNULL(EmpTimes_2.Hrs, 0) AS Hrs, ISNULL(EmpTimes_1.PayableHrs, 0) + ISNULL(EmpTimes_2.PayableHrs, 0) AS PayableHrs, 
                      ISNULL(EmpTimes_3.Hrs, 0) AS OTHrs, dbo.VEmpDailyEntries.TotalEntries, dbo.VEmpDailyEntries.SecondEntryID, dbo.VEmpDailyEntries.OverTimeEntryID, 
                      dbo.VEmpSettings.LunchInTime, dbo.VEmpSettings.LunchOutTime,AttendanceSheet.LateHours
FROM         dbo.VEmpDailyEntries INNER JOIN
                      dbo.Employees ON dbo.VEmpDailyEntries.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                      dbo.VEmpSettings ON dbo.Employees.empid = dbo.VEmpSettings.empid LEFT OUTER JOIN
                      dbo.VEmpPayableHrs EmpTimes_1 ON dbo.VEmpDailyEntries.FirstEntryID = EmpTimes_1.EntryID LEFT OUTER JOIN
                      dbo.VEmpPayableHrs EmpTimes_2 ON dbo.VEmpDailyEntries.SecondEntryID = EmpTimes_2.EntryID LEFT OUTER JOIN
                      dbo.VEmpPayableHrsOverTime EmpTimes_3 ON dbo.VEmpDailyEntries.OverTimeEntryID = EmpTimes_3.EntryID
					  INNER JOIN AttendanceSheet ON EmpTimes_1.EmpID=AttendanceSheet.EmpID AND EmpTimes_1.DT=AttendanceSheet.DT
ORDER BY dbo.Employees.empid

GO


