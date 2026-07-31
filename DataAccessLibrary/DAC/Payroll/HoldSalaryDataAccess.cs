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
    public interface IHoldSalaryDataAccess
    {
        Task<HoldSalarySummaryModel> GetEmployeeHoldSalaryInfoAsync(string empId, int year, int month);
        Task<List<HoldSalaryListItemModel>> GetHoldSalariesAsync(string? deptId, int year, int month);
        Task<bool> SaveHoldSalaryAsync(HoldSalaryDto dto, string userName, string machineName);
        Task<bool> DeleteHoldSalaryAsync(int entryId);
    }

    public class HoldSalaryDataAccess : IHoldSalaryDataAccess
    {
        private readonly string _connectionString;

        public HoldSalaryDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<HoldSalarySummaryModel> GetEmployeeHoldSalaryInfoAsync(string empId, int year, int month)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            var firstOfMonth = new DateTime(year, month, 1);
            var monthEnd = new DateTime(year, month, DateTime.DaysInMonth(year, month));

            const string sql = @"
                SELECT 
                    e.EmpID,
                    ISNULL(e.name, '') AS EmpName,
                    e.DeptID,
                    ISNULL(d.name, '') AS DeptName,
                    ISNULL((SELECT TOP 1 (ISNULL(m.Balance, 0) + ISNULL(m.HoldSalaryAmt, 0))
                            FROM MonthlySalaries m
                            WHERE m.EmpID = @EmpID AND m.DT = @MonthEndDT), 0) AS AvailableBalance,
                    ISNULL((SELECT TOP 1 ISNULL(h.HoldAmt, 0)
                            FROM EmpHoldSalary h
                            WHERE h.EmpID = @EmpID AND h.DT = @FirstOfMonth), 0) AS ExistingHoldAmt,
                    ISNULL((SELECT TOP 1 ISNULL(h.Description, '')
                            FROM EmpHoldSalary h
                            WHERE h.EmpID = @EmpID AND h.DT = @FirstOfMonth), '') AS ExistingDescription
                FROM VEmp e
                LEFT JOIN Departments d ON e.DeptID = d.DeptID
                WHERE e.EmpID = @EmpID";

            var result = await db.QueryFirstOrDefaultAsync<HoldSalarySummaryModel>(sql, new
            {
                EmpID = empId,
                FirstOfMonth = firstOfMonth,
                MonthEndDT = monthEnd
            });

            return result ?? new HoldSalarySummaryModel { EmpID = empId };
        }

        public async Task<List<HoldSalaryListItemModel>> GetHoldSalariesAsync(string? deptId, int year, int month)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                SELECT 
                    h.EntryID,
                    h.EmpID,
                    ISNULL(e.name, '') AS EmpName,
                    e.DeptID,
                    ISNULL(d.name, '') AS DeptName,
                    h.DT,
                    ISNULL(h.HoldAmt, 0) AS HoldAmt,
                    ISNULL(h.Description, '') AS Description,
                    h.DTEntry,
                    ISNULL(h.UserName, '') AS UserName,
                    ISNULL(h.MachineName, '') AS MachineName
                FROM EmpHoldSalary h
                LEFT JOIN VEmp e ON h.EmpID = e.EmpID
                LEFT JOIN Departments d ON e.DeptID = d.DeptID
                WHERE MONTH(h.DT) = @Month AND YEAR(h.DT) = @Year
                  AND (@DeptID IS NULL OR @DeptID = '0' OR e.DeptID = @DeptID)
                ORDER BY h.EntryID DESC";

            var list = await db.QueryAsync<HoldSalaryListItemModel>(sql, new
            {
                Month = month,
                Year = year,
                DeptID = string.IsNullOrWhiteSpace(deptId) ? null : deptId
            });

            return list.ToList();
        }

        public async Task<bool> SaveHoldSalaryAsync(HoldSalaryDto dto, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            var firstOfMonth = new DateTime(dto.Year, dto.Month, 1);
            var monthEnd = new DateTime(dto.Year, dto.Month, DateTime.DaysInMonth(dto.Year, dto.Month));

            const string sql = @"
                IF EXISTS (SELECT 1 FROM EmpHoldSalary WHERE EmpID = @EmpID AND DT = @FirstOfMonth)
                BEGIN
                    UPDATE EmpHoldSalary
                    SET HoldAmt = @HoldAmt,
                        Description = @Description,
                        UserName = @UserName,
                        MachineName = @MachineName,
                        DTEntry = GETDATE()
                    WHERE EmpID = @EmpID AND DT = @FirstOfMonth;
                END
                ELSE
                BEGIN
                    INSERT INTO EmpHoldSalary(EmpID, DT, HoldAmt, Description, UserName, MachineName, DTEntry)
                    VALUES(@EmpID, @FirstOfMonth, @HoldAmt, @Description, @UserName, @MachineName, GETDATE());
                END

                UPDATE MonthlySalaries
                SET HoldSalaryAmt = @HoldAmt
                WHERE EmpID = @EmpID AND DT = @MonthEndDT;";

            await db.ExecuteAsync(sql, new
            {
                EmpID = dto.EmpID,
                FirstOfMonth = firstOfMonth,
                MonthEndDT = monthEnd,
                HoldAmt = dto.HoldAmt,
                Description = dto.Description,
                UserName = userName,
                MachineName = machineName
            });

            return true;
        }

        public async Task<bool> DeleteHoldSalaryAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string getInfoSql = @"SELECT TOP 1 EmpID, DT FROM EmpHoldSalary WHERE EntryID = @EntryID";
            var info = await db.QueryFirstOrDefaultAsync<HoldSalaryListItemModel>(getInfoSql, new { EntryID = entryId });

            if (info != null && !string.IsNullOrEmpty(info.EmpID))
            {
                var monthEnd = new DateTime(info.DT.Year, info.DT.Month, DateTime.DaysInMonth(info.DT.Year, info.DT.Month));
                const string deleteSql = @"
                    DELETE FROM EmpHoldSalary WHERE EntryID = @EntryID;
                    UPDATE MonthlySalaries SET HoldSalaryAmt = 0 WHERE EmpID = @EmpID AND DT = @MonthEndDT;";

                await db.ExecuteAsync(deleteSql, new
                {
                    EntryID = entryId,
                    EmpID = info.EmpID,
                    MonthEndDT = monthEnd
                });
            }
            else
            {
                await db.ExecuteAsync("DELETE FROM EmpHoldSalary WHERE EntryID = @EntryID", new { EntryID = entryId });
            }

            return true;
        }
    }
}
