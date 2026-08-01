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
    public interface ISocialSecurityDataAccess
    {
        Task<SocialSecuritySheetModel> GetSocialSecuritySheetAsync(SocialSecurityFilterDto filter);
        Task<SocialSecuritySheetModel> GenerateSocialSecuritySheetAsync(SocialSecurityFilterDto filter);
        Task<bool> SaveSocialSecuritySheetAsync(SaveSocialSecurityDto dto);
    }

    public class SocialSecurityDataAccess : ISocialSecurityDataAccess
    {
        private readonly string _connectionString;

        public SocialSecurityDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<SocialSecuritySheetModel> GetSocialSecuritySheetAsync(SocialSecurityFilterDto filter)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string? deptIdParam = string.IsNullOrWhiteSpace(filter.DeptID) || filter.DeptID == "0"
                ? null
                : filter.DeptID;

            const string checkSql = @"
                SELECT COUNT(1)
                FROM SSMonthly s
                INNER JOIN Employees e ON s.EmpID = e.EmpID
                WHERE MONTH(s.DT) = @Month AND YEAR(s.DT) = @Year
                  AND (@DeptID IS NULL OR e.DeptID = @DeptID)";

            int count = await db.ExecuteScalarAsync<int>(checkSql, new
            {
                filter.Month,
                filter.Year,
                DeptID = deptIdParam
            });

            var sheet = new SocialSecuritySheetModel
            {
                Month = filter.Month,
                Year = filter.Year,
                DeptID = filter.DeptID,
                IsSaved = count > 0,
                Rows = new List<SocialSecurityRowModel>()
            };

            if (count > 0)
            {
                const string querySql = @"
                    SELECT 
                        v.empid AS EmpID,
                        LTRIM(RTRIM(ISNULL(v.name, ''))) AS EmpName,
                        LTRIM(RTRIM(ISNULL(v.fname, ''))) AS FName,
                        ISNULL(v.SocialNo, '') AS SocialNo,
                        ISNULL(v.Designation, '') AS Designation,
                        CASE WHEN v.EmpStatus = 1 OR v.EmpStatus IS NULL THEN 'Permanent' ELSE 'Temporary' END AS EmpStatus,
                        ISNULL(v.Salary, 0) AS Salary,
                        ISNULL(v.DaysWorked, 0) AS DaysWorked,
                        ISNULL(v.NetSalary, 0) AS NetSalary,
                        ISNULL(v.SSAmt, 0) AS SSAmt
                    FROM VSocialSecurity v
                    WHERE MONTH(v.DT) = @Month AND YEAR(v.DT) = @Year
                      AND (@DeptID IS NULL OR v.deptid = @DeptID)
                    ORDER BY v.empid";

                var rows = (await db.QueryAsync<SocialSecurityRowModel>(querySql, new
                {
                    filter.Month,
                    filter.Year,
                    DeptID = deptIdParam
                })).ToList();

                foreach (var r in rows)
                {
                    r.IsSelected = true;
                }
                sheet.Rows = rows;
            }

            return sheet;
        }

        public async Task<SocialSecuritySheetModel> GenerateSocialSecuritySheetAsync(SocialSecurityFilterDto filter)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            int daysInMonth = DateTime.DaysInMonth(filter.Year, filter.Month);
            DateTime parDate = new DateTime(filter.Year, filter.Month, daysInMonth);

            string? deptIdParam = string.IsNullOrWhiteSpace(filter.DeptID) || filter.DeptID == "0"
                ? null
                : filter.DeptID;

            var spParams = new DynamicParameters();
            spParams.Add("@ParDate", parDate);
            spParams.Add("@DeptID", deptIdParam);

            var rawRows = await db.QueryAsync("SP_SSMonthly", spParams, commandType: CommandType.StoredProcedure);

            var sheet = new SocialSecuritySheetModel
            {
                Month = filter.Month,
                Year = filter.Year,
                DeptID = filter.DeptID,
                IsSaved = false,
                Rows = new List<SocialSecurityRowModel>()
            };

            foreach (var r in rawRows)
            {
                var dict = (IDictionary<string, object>)r;

                string empId = Convert.ToString(GetDictValue(dict, "EmpID", "empid")) ?? string.Empty;
                string name = Convert.ToString(GetDictValue(dict, "Name", "name", "EmpName")) ?? string.Empty;
                string fname = Convert.ToString(GetDictValue(dict, "FName", "fname")) ?? string.Empty;
                string socialNo = Convert.ToString(GetDictValue(dict, "SocialNo", "socialno")) ?? string.Empty;
                string designation = Convert.ToString(GetDictValue(dict, "Designation", "designation")) ?? string.Empty;
                string empStatusRaw = Convert.ToString(GetDictValue(dict, "EmpStatus", "empstatus")) ?? "Permanent";

                decimal salary = Convert.ToDecimal(GetDictValue(dict, "EmpSalary", "Salary", "salary") ?? 0);
                int daysWorked = Convert.ToInt32(GetDictValue(dict, "SDays", "DaysWorked", "daysworked") ?? 0);
                decimal netSalary = Convert.ToDecimal(GetDictValue(dict, "SalAmt", "NetSalary", "netsalary") ?? 0);
                decimal ssAmt = Convert.ToDecimal(GetDictValue(dict, "SocialAmt", "SSAmt", "ssamt") ?? 0);

                if (!string.IsNullOrWhiteSpace(empId))
                {
                    sheet.Rows.Add(new SocialSecurityRowModel
                    {
                        EmpID = empId,
                        EmpName = name,
                        FName = fname,
                        SocialNo = socialNo,
                        Designation = designation,
                        EmpStatus = string.IsNullOrWhiteSpace(empStatusRaw) ? "Permanent" : empStatusRaw,
                        Salary = salary,
                        DaysWorked = daysWorked,
                        NetSalary = netSalary,
                        SSAmt = ssAmt,
                        IsSelected = true
                    });
                }
            }

            return sheet;
        }

        public async Task<bool> SaveSocialSecuritySheetAsync(SaveSocialSecurityDto dto)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            if (db.State != ConnectionState.Open)
            {
                db.Open();
            }

            int daysInMonth = DateTime.DaysInMonth(dto.Year, dto.Month);
            DateTime dt = new DateTime(dto.Year, dto.Month, daysInMonth);

            string? deptIdParam = string.IsNullOrWhiteSpace(dto.DeptID) || dto.DeptID == "0"
                ? null
                : dto.DeptID;

            using var trans = db.BeginTransaction();
            try
            {
                const string deleteSql = @"
                    DELETE FROM SSMonthly 
                    WHERE MONTH(DT) = @Month AND YEAR(DT) = @Year
                      AND (@DeptID IS NULL OR EmpID IN (SELECT EmpID FROM Employees WHERE DeptID = @DeptID))";

                await db.ExecuteAsync(deleteSql, new
                {
                    dto.Month,
                    dto.Year,
                    DeptID = deptIdParam
                }, trans);

                const string insertSql = @"
                    INSERT INTO SSMonthly (EmpID, DT, EmpStatus, Salary, DaysWorked, NetSalary, SSAmt)
                    VALUES (@EmpID, @DT, @EmpStatusBit, @Salary, @DaysWorked, @NetSalary, @SSAmt)";

                foreach (var row in dto.Rows.Where(r => r.IsSelected && !string.IsNullOrWhiteSpace(r.EmpID)))
                {
                    int empStatusBit = string.Equals(row.EmpStatus, "Permanent", StringComparison.OrdinalIgnoreCase) || row.EmpStatus == "1" ? 1 : 0;

                    await db.ExecuteAsync(insertSql, new
                    {
                        EmpID = row.EmpID,
                        DT = dt,
                        EmpStatusBit = empStatusBit,
                        Salary = row.Salary,
                        DaysWorked = row.DaysWorked,
                        NetSalary = row.NetSalary,
                        SSAmt = row.SSAmt
                    }, trans);
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

        private static object? GetDictValue(IDictionary<string, object> dict, params string[] keys)
        {
            foreach (var key in keys)
            {
                var match = dict.FirstOrDefault(k => string.Equals(k.Key, key, StringComparison.OrdinalIgnoreCase));
                if (match.Key != null && match.Value != null && match.Value != DBNull.Value)
                {
                    return match.Value;
                }
            }
            return null;
        }
    }
}
