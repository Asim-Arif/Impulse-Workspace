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
    public class EmployeeTransferDataAccess : IEmployeeTransferDataAccess
    {
        private readonly string _connectionString;

        public EmployeeTransferDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<DepartmentListItemModel>> GetSourceDepartmentsAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DeptID, Name, ISNULL(Active, 1) AS Active
                    FROM Departments
                    WHERE ISNULL(Active, 1) = 1
                    ORDER BY Name";
                return (await db.QueryAsync<DepartmentListItemModel>(sql)).AsList();
            }
        }

        public async Task<List<DepartmentListItemModel>> GetTargetDepartmentsAsync(string sourceDeptId)
        {
            if (string.IsNullOrWhiteSpace(sourceDeptId))
                return new List<DepartmentListItemModel>();

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string tempDeptSql = "SELECT ISNULL(TempDept, 0) FROM Departments WHERE DeptID = @DeptID";
                bool tempDept = await db.ExecuteScalarAsync<bool>(tempDeptSql, new { DeptID = sourceDeptId });

                string sql = @"
                    SELECT DeptID, Name, ISNULL(Active, 1) AS Active
                    FROM Departments
                    WHERE ISNULL(Active, 1) = 1
                      AND ISNULL(TempDept, 0) = @TempDept
                      AND DeptID <> @SourceDeptID
                    ORDER BY Name";

                return (await db.QueryAsync<DepartmentListItemModel>(sql, new { TempDept = tempDept ? 1 : 0, SourceDeptID = sourceDeptId })).AsList();
            }
        }

        public async Task<List<EmployeeListItemModel>> GetEmployeesByDepartmentAsync(string deptId)
        {
            if (string.IsNullOrWhiteSpace(deptId))
                return new List<EmployeeListItemModel>();

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EmpID, Name, ISNULL(FName, '') AS FName, ISNULL(Designation, '') AS Designation,
                           ISNULL(StartingSalary, 0) AS StartingSalary, ISNULL(Active, 1) AS Active, DeptID
                    FROM Employees
                    WHERE DeptID = @DeptID
                      AND ISNULL(Active, 1) = 1
                    ORDER BY Name";

                return (await db.QueryAsync<EmployeeListItemModel>(sql, new { DeptID = deptId })).AsList();
            }
        }

        public async Task<bool> SaveTransferAsync(EmployeeTransferDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.EmpID) || string.IsNullOrWhiteSpace(dto.TransferToDeptID))
                return false;

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        string insertSql = @"
                            INSERT INTO EmpTransfers (EmpID, TransferDT, TransferFromDeptID, TransferToDeptID, EntryDescription, UserName, MachineName, EntryDT)
                            VALUES (@EmpID, @TransferDT, @TransferFromDeptID, @TransferToDeptID, @EntryDescription, @UserName, 'WEB', GETDATE())";

                        await db.ExecuteAsync(insertSql, new
                        {
                            EmpID = dto.EmpID,
                            TransferDT = dto.TransferDT.Date,
                            TransferFromDeptID = dto.TransferFromDeptID,
                            TransferToDeptID = dto.TransferToDeptID,
                            EntryDescription = dto.EntryDescription,
                            UserName = dto.UserName
                        }, trans);

                        string updateEmpSql = @"
                            UPDATE Employees
                            SET DeptID = @TransferToDeptID
                            WHERE EmpID = @EmpID";

                        await db.ExecuteAsync(updateEmpSql, new
                        {
                            TransferToDeptID = dto.TransferToDeptID,
                            EmpID = dto.EmpID
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

        public async Task<List<EmployeeTransferHistoryRow>> GetRecentTransfersAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT TOP 50
                           t.EmpID, ISNULL(e.Name, '') AS EmpName, ISNULL(e.FName, '') AS FName, ISNULL(e.Designation, '') AS Designation,
                           ISNULL(fd.Name, t.TransferFromDeptID) AS FromDeptName,
                           ISNULL(td.Name, t.TransferToDeptID) AS ToDeptName,
                           t.TransferDT, ISNULL(t.EntryDescription, '') AS EntryDescription, ISNULL(t.UserName, '') AS UserName
                    FROM EmpTransfers t
                    LEFT JOIN Employees e ON t.EmpID = e.EmpID
                    LEFT JOIN Departments fd ON t.TransferFromDeptID = fd.DeptID
                    LEFT JOIN Departments td ON t.TransferToDeptID = td.DeptID
                    ORDER BY t.TransferDT DESC, t.EntryDT DESC";

                return (await db.QueryAsync<EmployeeTransferHistoryRow>(sql)).AsList();
            }
        }
    }
}
