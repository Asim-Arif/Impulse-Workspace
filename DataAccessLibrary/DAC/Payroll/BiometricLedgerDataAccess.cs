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
    public class BiometricLedgerDataAccess : IBiometricLedgerDataAccess
    {
        private readonly string _connectionString;

        public BiometricLedgerDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        // ── Dept Summary ────────────────────────────────────────────────────────────

        public async Task<List<BiometricDeptSummaryRow>> GetDeptSummaryAsync(DateTime date, string? deptId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var parameters = new DynamicParameters();
                parameters.Add("@DT",     date,   DbType.DateTime);
                parameters.Add("@DeptID", string.IsNullOrWhiteSpace(deptId) ? null : deptId, DbType.String, size: 50);

                var results = await db.QueryAsync<BiometricDeptSummaryRow>(
                    "SP_DeptSettings",
                    parameters,
                    commandType: CommandType.StoredProcedure,
                    commandTimeout: 120);

                return results.AsList();
            }
        }

        // ── Employee Attendance ─────────────────────────────────────────────────────

        public async Task<List<BiometricEmployeeRow>> GetEmployeeAttendanceAsync(DateTime date, string? deptId, string? empId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // Build WHERE clause for VEmpTimes1 subquery exactly as legacy RefreshLedger does
                string cond = $"WHERE (DT='{date:dd-MMM-yyyy}' OR ISNULL(DT,'')='')";

                if (string.IsNullOrWhiteSpace(deptId))
                {
                    // All depts
                    if (!string.IsNullOrWhiteSpace(empId))
                        cond += $" AND EmpID='{empId.Replace("'", "''")}'";
                }
                else
                {
                    // Specific dept
                    if (string.IsNullOrWhiteSpace(empId))
                        cond += $" AND DeptID='{deptId.Replace("'", "''")}'";
                    else
                        cond += $" AND EmpID='{empId.Replace("'", "''")}'";
                }

                string sql = $@"
                    SELECT 
                        RTRIM(LTRIM(ISNULL(Employees.DeptID,'')))  AS DeptID,
                        RTRIM(LTRIM(ISNULL(Employees.EmpID,'')))   AS EmpID,
                        Employees.Name               AS Name,
                        ISNULL(EmpTimes.EntryID, 0)  AS EntryID,
                        EmpTimes.DT                  AS DT,
                        ISNULL(EmpTimes.InTime,'')   AS InTime,
                        ISNULL(EmpTimes.OutTime,'')  AS OutTime,
                        ISNULL(EmpTimes.Hrs, 0)      AS Hrs,
                        ISNULL(EmpTimes.lateHours,0) AS LateHours,
                        ISNULL(EmpTimes.ExtraHrs,0)  AS ExtraHrs
                    FROM (
                        SELECT * FROM dbo.VEmpTimes1 {cond}
                    ) AS EmpTimes
                    RIGHT OUTER JOIN dbo.VEmp AS Employees
                        ON EmpTimes.EmpID = Employees.EmpID
                    WHERE Employees.Active = 1
                    ORDER BY Employees.DeptID, Employees.EmpID";

                var results = await db.QueryAsync<BiometricEmployeeRow>(sql, commandTimeout: 120);
                return results.AsList();
            }
        }

        // ── IFace Machines ─────────────────────────────────────────────────────────

        public async Task<List<BiometricMachineItem>> GetIFaceMachinesAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, ISNULL(IPAddress,'') AS IPAddress, 
                           ISNULL(Description,'') AS Description,
                           ISNULL(MachineNo, 1)   AS MachineNo,
                           ISNULL(MachineType, 3) AS MachineType
                    FROM AttendanceMachine
                    WHERE MachineType = 3
                    ORDER BY EntryID";

                var results = await db.QueryAsync<BiometricMachineItem>(sql);
                return results.AsList();
            }
        }

        // ── Report Staging Stored Procedures ──────────────────────────────────────

        public async Task ExecutePrintDailyAbsentLeaveDetailAsync(DateTime date)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var parameters = new DynamicParameters();
                parameters.Add("@DT", date, DbType.DateTime);

                await db.ExecuteAsync(
                    "SP_PrintDailyAbsentLeaveDetail",
                    parameters,
                    commandType: CommandType.StoredProcedure,
                    commandTimeout: 120);
            }
        }

        public async Task ExecutePrintDailyAbsentLeaveSummaryAsync(DateTime date)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var parameters = new DynamicParameters();
                parameters.Add("@DT", date, DbType.DateTime);

                await db.ExecuteAsync(
                    "SP_PrintDailyAbsentLeaveSummary",
                    parameters,
                    commandType: CommandType.StoredProcedure,
                    commandTimeout: 120);
            }
        }
    }
}
