USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VPaySlip]    Script Date: 10/5/2026 1:51:15 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VPaySlip]
AS
SELECT        dbo.Departments.deptid, dbo.Departments.name AS DeptName, dbo.Employees.empid, dbo.Employees.name, dbo.Employees.Rel, dbo.Employees.fname, dbo.PrintSalary.BSal, dbo.PrintSalary.Rate, dbo.PrintSalary.ADays, 
                         dbo.PrintSalary.AAmt, dbo.PrintSalary.SDays, dbo.PrintSalary.SAmt, dbo.PrintSalary.OHrs, dbo.PrintSalary.OAmt, dbo.PrintSalary.LHrs, dbo.PrintSalary.LAmt, dbo.PrintSalary.Total, dbo.PrintSalary.Tax, dbo.PrintSalary.NetTtl, 
                         dbo.PrintSalary.Paid, dbo.PrintSalary.Balance, dbo.PrintSalary.DT, dbo.Employees.Salary, dbo.Employees.PProfit, dbo.Employees.StartingSalary, dbo.PrintSalary.LongTerm, dbo.Employees.Designation, 
                         dbo.PrintSalary.AdvSal, dbo.Employees.EmpType, dbo.Employees.Status, dbo.PrintSalary.PrevLTLoan, dbo.PrintSalary.AdvPer, dbo.PrintSalary.AAllow, dbo.PrintSalary.AAllowAmt, dbo.PrintSalary.Leaves, 
                         dbo.PrintSalary.LeaveAmt, dbo.PrintSalary.Fine, dbo.PrintSalary.Bonus, dbo.PrintSalary.SocialAmt, dbo.Employees.NTN, dbo.Employees.BankPymt, dbo.PrintSalary.CasualLeaves, dbo.PrintSalary.SickLeaves, 
                         dbo.PrintSalary.AnnualLeaves, dbo.PrintSalary.CompensatoryLeaves, dbo.PrintSalary.MaternityLeaves, dbo.PrintSalary.WPLeaves, dbo.Employees.sex, dbo.Employees.Lunch, dbo.PrintSalary.EOBI, dbo.Employees.EmpIDOld, 
                         dbo.PrintSalary.SundayOTHrs, dbo.PrintSalary.SundayOTRate, dbo.PrintSalary.FixAllowance, dbo.Employees.EOBIAmt, dbo.Employees.SocialSecurityAmt, dbo.Employees.EmpLunchAmt, dbo.PrintSalary.HoldSalaryAmt, 
                         dbo.PrintSalary.PresentDays, dbo.PrintSalary.LeaveDays, dbo.PrintSalary.LateComingHrs, dbo.PrintSalary.ShortHrs, dbo.PrintSalary.AmtPaid, dbo.PrintSalary.GPHrs, dbo.PrintSalary.GPHrsAmt, dbo.PrintSalary.OTDinnerCount, 
                         dbo.PrintSalary.OTDinnerAmount, dbo.PrintSalary.DedOnePercent, dbo.PrintSalary.PerformanceDedAmt, dbo.PrintSalary.RejectionDedAmt, dbo.PrintSalary.ShortHrsAmt, dbo.PrintSalary.ZeroAbsentBonus, 
                         dbo.VEmpSettings.OTRate, dbo.PrintSalary.OTHrs_Original, dbo.PrintSalary.LateHrs_Original, dbo.PrintSalary.OTHrs_Net, dbo.PrintSalary.LateHrs_Net
FROM            dbo.Departments INNER JOIN
                         dbo.PrintSalary INNER JOIN
                         dbo.Employees ON dbo.PrintSalary.EmpID = dbo.Employees.empid ON dbo.Departments.deptid = dbo.PrintSalary.DeptID INNER JOIN
                         dbo.VEmpSettings ON dbo.Employees.empid = dbo.VEmpSettings.empid
GO


