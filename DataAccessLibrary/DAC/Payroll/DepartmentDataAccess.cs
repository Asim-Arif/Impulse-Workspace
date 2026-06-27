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
    public class DepartmentDataAccess : IDepartmentDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<DepartmentDataAccess> _logger;

        public DepartmentDataAccess(IConfiguration configuration, ILogger<DepartmentDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<DepartmentModel>> GetDepartmentsListAsync(bool showInactive)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"
                        SELECT deptid, Name, Description, Active, Color, Dept_ID, TempDept, 
                               SalaryAccNo, AccruedAccNo, AccruedFoodAccNo, AccruedUnionAccNo, 
                               SchAccNo, AccruedEOBIAccNo, TaxAccNo, OnePercentAccNo 
                        FROM Departments";

                    if (!showInactive)
                    {
                        sql += " WHERE Active = 1";
                    }

                    sql += " ORDER BY deptid";

                    var result = await db.QueryAsync<DepartmentModel>(sql);
                    return result.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching department list");
                throw;
            }
        }

        public async Task<DepartmentModel?> GetDepartmentByIdAsync(string deptId)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"
                        SELECT deptid, Name, Description, Active, Color, Dept_ID, TempDept, 
                               SalaryAccNo, AccruedAccNo, AccruedFoodAccNo, AccruedUnionAccNo, 
                               SchAccNo, AccruedEOBIAccNo, TaxAccNo, OnePercentAccNo 
                        FROM Departments 
                        WHERE deptid = @DeptID";

                    var result = await db.QueryFirstOrDefaultAsync<DepartmentModel>(sql, new { DeptID = deptId });
                    return result;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching department {DeptID}", deptId);
                throw;
            }
        }

        public async Task SaveDepartmentAsync(DepartmentModel department, bool isAdd)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    if (isAdd)
                    {
                        string sql = @"
                            INSERT INTO Departments (
                                DeptID, Name, Description, Active, Color, TempDept, 
                                SalaryAccNo, AccruedAccNo, AccruedFoodAccNo, AccruedUnionAccNo, 
                                SchAccNo, AccruedEOBIAccNo, TaxAccNo, OnePercentAccNo
                            ) VALUES (
                                @DeptID, @Name, @Description, @Active, @Color, @TempDept, 
                                @SalaryAccNo, @AccruedAccNo, @AccruedFoodAccNo, @AccruedUnionAccNo, 
                                @SchAccNo, @AccruedEOBIAccNo, @TaxAccNo, @OnePercentAccNo
                            )";
                        await db.ExecuteAsync(sql, department);
                    }
                    else
                    {
                        string sql = @"
                            UPDATE Departments SET 
                                Name = @Name, 
                                Description = @Description, 
                                Color = @Color, 
                                TempDept = @TempDept, 
                                SalaryAccNo = @SalaryAccNo, 
                                AccruedAccNo = @AccruedAccNo, 
                                AccruedFoodAccNo = @AccruedFoodAccNo, 
                                AccruedUnionAccNo = @AccruedUnionAccNo, 
                                SchAccNo = @SchAccNo, 
                                AccruedEOBIAccNo = @AccruedEOBIAccNo, 
                                TaxAccNo = @TaxAccNo, 
                                OnePercentAccNo = @OnePercentAccNo 
                            WHERE DeptID = @DeptID";
                        await db.ExecuteAsync(sql, department);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving department {DeptID}", department.DeptID);
                throw;
            }
        }

        public async Task<string> GetNextDeptIdAsync()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "SELECT FORMAT(ISNULL(MAX(CAST(deptid AS INT)), 0) + 1, '0000') FROM Departments";
                    var nextId = await db.ExecuteScalarAsync<string>(sql);
                    return nextId ?? "0001";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting next department ID");
                throw;
            }
        }

        public async Task SetDepartmentActiveStateAsync(string deptId, bool active)
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
                            // Update department active state
                            string updateDeptSql = "UPDATE Departments SET Active = @Active WHERE DeptID = @DeptID";
                            await db.ExecuteAsync(updateDeptSql, new { DeptID = deptId, Active = active ? 1 : 0 }, transaction: trans);

                            // Update employee active states for that department, matching VB6 logic
                            string updateEmpSql;
                            if (active)
                            {
                                updateEmpSql = "UPDATE Employees SET Active = 1, DiscontinuedOn = NULL WHERE DeptID = @DeptID";
                                await db.ExecuteAsync(updateEmpSql, new { DeptID = deptId }, transaction: trans);
                            }
                            else
                            {
                                updateEmpSql = "UPDATE Employees SET Active = 0, DiscontinuedOn = @DT WHERE DeptID = @DeptID";
                                await db.ExecuteAsync(updateEmpSql, new { DeptID = deptId, DT = DateTime.Today }, transaction: trans);
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
                _logger.LogError(ex, "Error toggling active state for department {DeptID} to {Active}", deptId, active);
                throw;
            }
        }
    }
}
