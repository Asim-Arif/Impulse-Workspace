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
    public class SalaryHistoryDataAccess : ISalaryHistoryDataAccess
    {
        private readonly string _connectionString;

        public SalaryHistoryDataAccess(IConfiguration config)
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

        public async Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var sb = new StringBuilder();
                sb.Append(@"
                    SELECT EmpID, Name, ISNULL(FName, '') AS FName, ISNULL(Designation, '') AS Designation,
                           ISNULL(StartingSalary, 0) AS StartingSalary, ISNULL(Active, 1) AS Active, DeptID
                    FROM Employees
                    WHERE ISNULL(Active, 1) = 1");

                if (!string.IsNullOrWhiteSpace(deptId) && deptId != "0")
                {
                    sb.Append(" AND DeptID = @DeptID");
                }

                sb.Append(" ORDER BY Name");

                var list = (await db.QueryAsync<EmployeeListItemModel>(sb.ToString(), new { DeptID = deptId })).AsList();
                list.Insert(0, new EmployeeListItemModel { EmpID = "0", Name = "<All Employees>", Active = true });
                return list;
            }
        }

        public async Task<List<SalaryHistoryRowModel>> GetSalaryHistoryAsync(SalaryHistoryFilterModel filter)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var sb = new StringBuilder();
                sb.Append(@"
                    SELECT DeptID, ISNULL(DeptName, '') AS DeptName, EmpID, ISNULL(Name, '') AS Name,
                           ISNULL(FName, '') AS FName, ISNULL(Rel, '') AS Rel, ISNULL(Designation, '') AS Designation,
                           DT, ISNULL(Description, '') AS Description, ISNULL(Salary, 0) AS Salary,
                           ISNULL(SalaryExt, 0) AS SalaryExt, PreviousSalary
                    FROM VSalaryLedger
                    WHERE CAST(DT AS DATE) BETWEEN @FromDate AND @ToDate
                      AND ISNULL(TempDept, 0) = 0
                      AND ISNULL(Active, 1) = 1");

                if (!string.IsNullOrWhiteSpace(filter.DeptID) && filter.DeptID != "0")
                {
                    sb.Append(" AND DeptID = @DeptID");
                }

                if (!string.IsNullOrWhiteSpace(filter.EmpID) && filter.EmpID != "0")
                {
                    sb.Append(" AND EmpID = @EmpID");
                }

                sb.Append(" ORDER BY DT ASC");

                var results = await db.QueryAsync<SalaryHistoryRowModel>(sb.ToString(), new
                {
                    FromDate = filter.FromDate.Date,
                    ToDate = filter.ToDate.Date,
                    DeptID = filter.DeptID,
                    EmpID = filter.EmpID
                });

                return results.AsList();
            }
        }

        public async Task<bool> SaveSalaryChangeAsync(ChangeSalaryRequestModel request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.EmpID) || request.EmpID == "0")
                return false;

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        // Get current salary
                        string prevSql = "SELECT ISNULL(StartingSalary, 0) FROM Employees WHERE EmpID = @EmpID";
                        double prevSalary = await db.ExecuteScalarAsync<double>(prevSql, new { EmpID = request.EmpID }, trans);

                        // Insert into SalaryLedger
                        string ledgerSql = @"
                            INSERT INTO SalaryLedger(EmpID, Salary, SalaryExt, DT, Description, UserName, MachineName, EntryDT, PreviousSalary)
                            VALUES(@EmpID, @Salary, @SalaryExt, @DT, @Description, @UserName, 'WEB', GETDATE(), @PreviousSalary)";

                        await db.ExecuteAsync(ledgerSql, new
                        {
                            EmpID = request.EmpID,
                            Salary = request.NewSalary,
                            SalaryExt = request.NewSalaryExt,
                            DT = request.EffectiveDate.Date,
                            Description = request.Description,
                            UserName = request.UserName,
                            PreviousSalary = (int)prevSalary
                        }, trans);

                        // Update Employees table StartingSalary
                        string updateEmpSql = "UPDATE Employees SET StartingSalary = @NewSalary WHERE EmpID = @EmpID";
                        await db.ExecuteAsync(updateEmpSql, new
                        {
                            NewSalary = request.NewSalary,
                            EmpID = request.EmpID
                        }, trans);

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

        public async Task PrepareAbsentListReportProcAsync(DateTime fromDate, DateTime toDate)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var p = new DynamicParameters();
                p.Add("@DTFrom", fromDate.Date);
                p.Add("@DTTo", toDate.Date);

                await db.ExecuteAsync("SP_PrintEmpAbsentList", p, commandType: CommandType.StoredProcedure);
            }
        }
    }
}
