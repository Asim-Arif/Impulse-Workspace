using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class OverTimeAuthDataAccess : IOverTimeAuthDataAccess
    {
        private readonly string _connectionString;

        public OverTimeAuthDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<DepartmentListItemModel>> GetDepartmentsAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DeptID, Name, ISNULL(TempDept, 0) AS TempDept
                    FROM Departments
                    WHERE ISNULL(TempDept, 0) = 0
                    ORDER BY Name";

                return (await db.QueryAsync<DepartmentListItemModel>(sql)).AsList();
            }
        }

        public async Task<List<OverTimeAuthRowDto>> GetOverTimeAuthListAsync(string deptId, DateTime date)
        {
            if (string.IsNullOrWhiteSpace(deptId)) return new List<OverTimeAuthRowDto>();

            DateTime attDate = date.Date;
            int monthDays = DateTime.DaysInMonth(attDate.Year, attDate.Month);

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // Check if Holiday or Sunday
                string holidaySql = "SELECT COUNT(1) FROM Holidays WHERE CAST(DT AS DATE) = @DT";
                int holidayCount = await db.ExecuteScalarAsync<int>(holidaySql, new { DT = attDate });
                bool isSundayOrHoliday = (attDate.DayOfWeek == DayOfWeek.Sunday) || (holidayCount > 0);

                string empSql;
                if (isSundayOrHoliday)
                {
                    empSql = @"
                        SELECT EmpID, DeptID, Name AS EmpName, ISNULL(FName, '') AS FName, ISNULL(Rel, '') AS Rel,
                               ISNULL(StartingSalary, 0) AS StartingSalary
                        FROM Employees
                        WHERE DeptID = @DeptID AND ISNULL(Active, 1) = 1
                        ORDER BY EmpID";
                }
                else
                {
                    empSql = @"
                        SELECT EmpID, DeptID, Name AS EmpName, ISNULL(FName, '') AS FName, ISNULL(Rel, '') AS Rel,
                               ISNULL(StartingSalary, 0) AS StartingSalary
                        FROM Employees
                        WHERE DeptID = @DeptID AND ISNULL(Active, 1) = 1
                          AND EmpID IN (SELECT EmpID FROM AttendanceSheet WHERE Attendance = 1 AND CAST(DT AS DATE) = @DT)
                        ORDER BY EmpID";
                }

                var empList = (await db.QueryAsync<OverTimeAuthRowDto>(empSql, new { DeptID = deptId, DT = attDate })).AsList();
                if (!empList.Any()) return empList;

                // Query existing EmpTimes payable hours
                string timesSql = @"
                    SELECT EmpID, ISNULL(SUM(PayableHrs), 0) AS PayableHrs
                    FROM EmpTimes
                    WHERE CAST(DT AS DATE) = @DT AND EmpID IN (SELECT EmpID FROM Employees WHERE DeptID = @DeptID)
                    GROUP BY EmpID";
                var timesDict = (await db.QueryAsync(timesSql, new { DT = attDate, DeptID = deptId }))
                                .ToDictionary(x => (string)x.EmpID, x => (double)x.PayableHrs);

                // Query existing EmpOTHrs authorization records
                string otSql = @"
                    SELECT EntryID, EmpID, ISNULL(OTHrs, 0) AS OTHrs
                    FROM EmpOTHrs
                    WHERE CAST(DT AS DATE) = @DT AND EmpID IN (SELECT EmpID FROM Employees WHERE DeptID = @DeptID)";
                var otDict = (await db.QueryAsync(otSql, new { DT = attDate, DeptID = deptId }))
                             .ToDictionary(x => (string)x.EmpID, x => new { EntryID = (int)x.EntryID, OTHrs = (double)x.OTHrs });

                foreach (var row in empList)
                {
                    // Rate formula: (StartingSalary / MonthDays / 8) * 2
                    if (monthDays > 0)
                    {
                        row.OtRate = Math.Round((row.StartingSalary / monthDays / 8m) * 2m, 4);
                    }

                    if (otDict.TryGetValue(row.EmpID, out var otVal))
                    {
                        row.EntryID = otVal.EntryID;
                        row.OtHrs = otVal.OTHrs;
                        row.IsAuthorized = true;
                    }
                    else if (timesDict.TryGetValue(row.EmpID, out var payableHrs))
                    {
                        row.OtHrs = payableHrs;
                        row.IsAuthorized = false;
                    }
                    else
                    {
                        row.OtHrs = 0;
                        row.IsAuthorized = false;
                    }
                }

                return empList;
            }
        }

        public async Task<bool> SaveOverTimeAuthAsync(string deptId, DateTime date, List<OverTimeAuthRowDto> rows)
        {
            if (rows == null) return false;

            DateTime attDate = date.Date;

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        foreach (var row in rows)
                        {
                            if (row.IsAuthorized)
                            {
                                string checkSql = "SELECT EntryID FROM EmpOTHrs WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                                int? existingEntryId = await db.ExecuteScalarAsync<int?>(checkSql, new { EmpID = row.EmpID, DT = attDate }, trans);

                                if (existingEntryId.HasValue)
                                {
                                    string updateSql = "UPDATE EmpOTHrs SET OTHrs = @OTHrs WHERE EntryID = @EntryID";
                                    await db.ExecuteAsync(updateSql, new { OTHrs = row.OtHrs, EntryID = existingEntryId.Value }, trans);
                                }
                                else
                                {
                                    string insertSql = "INSERT INTO EmpOTHrs (EmpID, DT, OTHrs) VALUES (@EmpID, @DT, @OTHrs)";
                                    await db.ExecuteAsync(insertSql, new { EmpID = row.EmpID, DT = attDate, OTHrs = row.OtHrs }, trans);
                                }
                            }
                            else
                            {
                                string delSql = "DELETE FROM EmpOTHrs WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                                await db.ExecuteAsync(delSql, new { EmpID = row.EmpID, DT = attDate }, trans);
                            }
                        }

                        trans.Commit();
                        return true;
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
