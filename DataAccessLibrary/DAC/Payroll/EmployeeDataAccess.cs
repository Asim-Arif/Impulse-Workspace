using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Payroll
{
    public class EmployeeDataAccess : IEmployeeDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<EmployeeDataAccess> _logger;

        public EmployeeDataAccess(IConfiguration configuration, ILogger<EmployeeDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<DepartmentListItemModel>> GetDepartmentsAsync(bool showInactive, string? selectedDeptId = null)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "SELECT deptid AS DeptID, name AS Name, Active, Color, TempDept FROM Departments WHERE IsNull(TempDept, 0) = 0";
                    
                    if (!showInactive)
                    {
                        sql += " AND Active = 1";
                    }

                    if (!string.IsNullOrEmpty(selectedDeptId) && selectedDeptId != "0")
                    {
                        sql += " AND deptid = @DeptID";
                    }

                    sql += " ORDER BY name";

                    var result = await db.QueryAsync<DepartmentListItemModel>(sql, new { DeptID = selectedDeptId });
                    return result.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching departments");
                throw;
            }
        }

        public async Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId, bool showInactive, bool tempEmpOnly, string? selectedEmpId = null)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "SELECT EmpID, Name, FName, Designation, StartingSalary, JoinDate, Active, TempEmp, Phone1, Phone2, DeptID, DeptName FROM VEmp WHERE DeptID = @DeptID";

                    if (!showInactive)
                    {
                        sql += " AND Active = 1";
                    }

                    if (tempEmpOnly)
                    {
                        sql += " AND TempEmp = 1";
                    }
                    else
                    {
                        sql += " AND TempEmp = 0";
                    }

                    if (!string.IsNullOrEmpty(selectedEmpId) && selectedEmpId != "0")
                    {
                        sql += " AND EmpID = @EmpID";
                    }

                    sql += " ORDER BY EmpID";

                    var result = await db.QueryAsync<EmployeeListItemModel>(sql, new { DeptID = deptId, EmpID = selectedEmpId });
                    return result.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching employees for department {DeptID}", deptId);
                throw;
            }
        }

        public async Task<byte[]?> GetEmployeePictureAsync(string empId)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "SELECT EmpPic FROM Employees WHERE EmpID = @EmpID";
                    var pic = await db.ExecuteScalarAsync<byte[]>(sql, new { EmpID = empId });
                    return pic;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching picture for employee {EmpID}", empId);
                throw;
            }
        }

        public async Task SetEmployeeActiveStateAsync(string empId, bool active, DateTime? date = null, string? remarks = null, string? userName = null, string? machineName = null)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    using (var trans = db.BeginTransaction())
                    {
                        try
                        {
                            if (!active)
                            {
                                // Deactivate employee and set DiscontinuedOn date
                                string updateSql = "UPDATE Employees SET Active = 0, DiscontinuedOn = @DT WHERE EmpID = @EmpID";
                                await db.ExecuteAsync(updateSql, new { EmpID = empId, DT = date ?? DateTime.Today }, transaction: trans);

                                // Insert inactive audit trail
                                string insertSql = @"
                                    INSERT INTO EmployeeInActiveDetails (EmpID, DT, InAciveRemarks, UserName, MachineName) 
                                    VALUES (@EmpID, @DT, @Remarks, @UserName, @MachineName)";
                                await db.ExecuteAsync(insertSql, new
                                {
                                    EmpID = empId,
                                    DT = date ?? DateTime.Today,
                                    Remarks = remarks ?? "Discontinued",
                                    UserName = userName ?? "System",
                                    MachineName = machineName ?? "Web"
                                }, transaction: trans);
                            }
                            else
                            {
                                // Re-activate employee, clear DiscontinuedOn
                                string updateSql = "UPDATE Employees SET Active = 1, DiscontinuedOn = NULL WHERE EmpID = @EmpID";
                                await db.ExecuteAsync(updateSql, new { EmpID = empId }, transaction: trans);
                            }

                            trans.Commit();
                        }
                        catch (Exception)
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating active status of employee {EmpID} to {Active}", empId, active);
                throw;
            }
        }
    }
}
