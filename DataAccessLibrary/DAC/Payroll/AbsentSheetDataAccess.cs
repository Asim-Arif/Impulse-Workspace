using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Payroll
{
    public interface IAbsentSheetDataAccess
    {
        Task<AbsentSheetResultModel> GetAbsentSheetAsync(AbsentSheetFilterDto filter);
    }

    public class AbsentSheetDataAccess : IAbsentSheetDataAccess
    {
        private readonly string _connectionString;

        public AbsentSheetDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<AbsentSheetResultModel> GetAbsentSheetAsync(AbsentSheetFilterDto filter)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string? deptIdParam = string.IsNullOrWhiteSpace(filter.DeptID) || filter.DeptID == "0"
                ? null
                : filter.DeptID;

            var spParams = new DynamicParameters();
            spParams.Add("@DTFrom", filter.FromDate.Date);
            spParams.Add("@DTTo", filter.ToDate.Date);
            spParams.Add("@DeptID", deptIdParam);

            // Execute SP_PrintEmpAbsentList to populate PrintEmpAbsentList table
            await db.ExecuteAsync("SP_PrintEmpAbsentList", spParams, commandType: CommandType.StoredProcedure);

            // Select results from PrintEmpAbsentList joined with Employees and Departments
            const string querySql = @"
                SELECT 
                    T1.EmpID,
                    T1.DT AS Date,
                    LTRIM(RTRIM(ISNULL(T2.Name, '') + ' ' + ISNULL(T2.Rel, '') + ' ' + ISNULL(T2.FName, ''))) AS EmpName,
                    ISNULL(T2.DeptID, '') AS DeptID,
                    ISNULL(d.Name, '') AS DeptName,
                    ISNULL(T2.Designation, '') AS Designation
                FROM PrintEmpAbsentList T1
                INNER JOIN Employees T2 ON T1.EmpID = T2.EmpID
                LEFT JOIN Departments d ON T2.DeptID = d.DeptID
                ORDER BY T1.DT, T1.EmpID";

            var rows = (await db.QueryAsync<AbsentSheetRowModel>(querySql)).ToList();

            return new AbsentSheetResultModel
            {
                Rows = rows
            };
        }
    }
}
