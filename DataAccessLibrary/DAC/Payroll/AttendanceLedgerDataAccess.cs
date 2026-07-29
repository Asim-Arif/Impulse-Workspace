using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class AttendanceLedgerDataAccess : IAttendanceLedgerDataAccess
    {
        private readonly string _connectionString;

        public AttendanceLedgerDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<DepartmentListItemModel>> GetDepartmentsAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DeptID, Name, ISNULL(Active, 1) AS Active
                    FROM Departments
                    WHERE ISNULL(Active, 1) = 1
                    ORDER BY Name";
                var list = (await db.QueryAsync<DepartmentListItemModel>(sql)).AsList();
                list.Insert(0, new DepartmentListItemModel { DeptID = "0", Name = "<All Departments>", Active = true });
                return list;
            }
        }

        public async Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId, DateTime firstDT, DateTime lastDT)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // Legacy EmpLedger: employees who had joined by month-end and were not
                // discontinued before month-start (department filter optional).
                var sb = new StringBuilder();
                sb.Append(@"
                    SELECT EmpID, ISNULL(Name, '') AS Name, ISNULL(FName, '') AS FName,
                           ISNULL(Designation, '') AS Designation, ISNULL(Active, 1) AS Active, DeptID
                    FROM Employees
                    WHERE JoinDate <= @LastDT AND (DiscontinuedOn > @FirstDT OR DiscontinuedOn IS NULL)");

                if (!string.IsNullOrWhiteSpace(deptId) && deptId != "0")
                {
                    sb.Append(" AND DeptID = @DeptID");
                }

                sb.Append(" ORDER BY Name");

                var list = (await db.QueryAsync<EmployeeListItemModel>(sb.ToString(), new
                {
                    DeptID = deptId,
                    FirstDT = firstDT.Date,
                    LastDT = lastDT.Date
                })).AsList();
                return list;
            }
        }

        public async Task<List<HolidayItemModel>> GetHolidaysAsync(int month, int year)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DAY(DT) AS Day, ISNULL(Description, '') AS Description
                    FROM Holidays
                    WHERE YEAR(DT) = @Year AND MONTH(DT) = @Month
                    ORDER BY DT";
                return (await db.QueryAsync<HolidayItemModel>(sql, new { Month = month, Year = year })).AsList();
            }
        }

        public async Task<List<AttendanceLedgerEntryModel>> GetLedgerEntriesAsync(AttendanceLedgerFilterModel filter)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DT, AttDay, Attendance,
                           ISNULL(CL, 0) AS CL, ISNULL(SL, 0) AS SL, ISNULL(AL, 0) AS AL,
                           ISNULL(CPL, 0) AS CPL, ISNULL(LWP, 0) AS LWP, ISNULL(ML, 0) AS ML,
                           ISNULL(SAL, 0) AS SAL, OTHrsOther
                    FROM VEmpLedger
                    WHERE YEAR(DT) = @Year AND MONTH(DT) = @Month AND EmpID = @EmpID
                    ORDER BY DT";
                return (await db.QueryAsync<AttendanceLedgerEntryModel>(sql, new
                {
                    filter.Year,
                    filter.Month,
                    filter.EmpID
                })).AsList();
            }
        }

        public async Task<List<AttendanceTimeModel>> GetTimeEntriesAsync(AttendanceLedgerFilterModel filter)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DT,
                           ISNULL(FirstInTime, '') AS FirstInTime, ISNULL(FirstOutTime, '') AS FirstOutTime,
                           ISNULL(SecondInTime, '') AS SecondInTime, ISNULL(SecondOutTime, '') AS SecondOutTime,
                           ISNULL(OTInTime, '') AS OTInTime, ISNULL(OTOutTime, '') AS OTOutTime,
                           ISNULL(Hrs, 0) AS Hrs, ISNULL(PayableHrs, 0) AS PayableHrs,
                           ISNULL(OTHrs, 0) AS OTHrs, ISNULL(LateHours, 0) AS LateHours
                    FROM VEmpTimes2
                    WHERE EmpID = @EmpID AND YEAR(DT) = @Year AND MONTH(DT) = @Month
                    ORDER BY DT, EmpID, FirstEntryID";
                return (await db.QueryAsync<AttendanceTimeModel>(sql, new
                {
                    filter.EmpID,
                    filter.Year,
                    filter.Month
                })).AsList();
            }
        }

        public async Task<Dictionary<DateTime, string>> GetShortLeavesAsync(AttendanceLedgerFilterModel filter)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DT, dbo.FShortLeaves(DT, @EmpID) AS ShortLeaves
                    FROM Leaves
                    WHERE EmpID = @EmpID AND MONTH(DT) = @Month AND YEAR(DT) = @Year
                    GROUP BY DT";
                var rows = await db.QueryAsync<(DateTime DT, string ShortLeaves)>(sql, new
                {
                    filter.EmpID,
                    filter.Month,
                    filter.Year
                });

                return rows
                    .Where(r => !string.IsNullOrWhiteSpace(r.ShortLeaves))
                    .GroupBy(r => r.DT.Date)
                    .ToDictionary(g => g.Key, g => g.First().ShortLeaves.Trim());
            }
        }

        public async Task PrepareEmpLedgerPrintAsync(List<AttendanceLedgerDayRow> rows, string empId)
        {
            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        await db.ExecuteAsync("DELETE FROM PrintEmpLedger", transaction: trans);

                        string insertSql = @"
                            INSERT INTO PrintEmpLedger
                                (DT, Status, InTime, OutTime, SecondInTime, SecondOutTime, Hrs,
                                 OTInTime, OTOutTime, LateHours, OverTime, ShortLeave, EmpID, PayableHrs, OTHrsOther)
                            VALUES
                                (@DT, @Status, @InTime, @OutTime, @SecondInTime, @SecondOutTime, @Hrs,
                                 @OTInTime, @OTOutTime, @LateHours, @OverTime, @ShortLeave, @EmpID, @PayableHrs, @OTHrsOther)";

                        foreach (var row in rows)
                        {
                            await db.ExecuteAsync(insertSql, new
                            {
                                DT = row.Date,
                                Status = row.Status,
                                InTime = row.FirstInTime,
                                OutTime = row.FirstOutTime,
                                row.SecondInTime,
                                row.SecondOutTime,
                                Hrs = row.ShowTimes ? AttendanceLedgerFormatter.GetHrsMin(row.Hrs) : string.Empty,
                                row.OTInTime,
                                row.OTOutTime,
                                LateHours = string.Empty,
                                OverTime = row.ShowTimes ? AttendanceLedgerFormatter.GetHrsMin(row.OTHrs, true) : string.Empty,
                                ShortLeave = row.ShortLeaves,
                                EmpID = empId,
                                PayableHrs = row.ShowTimes ? AttendanceLedgerFormatter.GetHrsMin(Math.Min(row.PayableHrs, 8), true) : string.Empty,
                                row.OTHrsOther
                            }, trans);
                        }

                        // Legacy: refresh OT / late hours from the datewise functions
                        string updateOtSql = @"
                            UPDATE T1 SET T1.OTHrs = T2.OTHrs, T1.LateHrs = T2.LateHrs
                            FROM PrintEmpLedger T1
                            INNER JOIN dbo.OverTime_Datewise_Fn((SELECT MAX(DT) FROM PrintEmpLedger), @EmpID) T2
                                ON T1.EmpID = T2.EmpID AND T1.DT = T2.DT";
                        await db.ExecuteAsync(updateOtSql, new { EmpID = empId }, trans);

                        string updateLateSql = @"
                            UPDATE T1 SET T1.LateHrs_Original = T2.LateComingHrs
                            FROM PrintEmpLedger T1
                            INNER JOIN dbo.LateComingHrs_Datewise_Fn((SELECT MAX(DT) FROM PrintEmpLedger)) T2
                                ON T1.EmpID = T2.EmpID AND T1.DT = T2.DT";
                        await db.ExecuteAsync(updateLateSql, new { EmpID = empId }, trans);

                        trans.Commit();
                    }
                    catch
                    {
                        trans.Rollback();
                        throw;
                    }
                }
            }
        }
    }
}
