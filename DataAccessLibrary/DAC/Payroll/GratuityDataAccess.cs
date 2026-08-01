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
    public interface IGratuityDataAccess
    {
        Task<GratuitySheetModel> GetGratuitySheetAsync(int year, string? deptId);
        Task<bool> SaveGratuitySheetAsync(SaveGratuityDto dto);
    }

    public class GratuityDataAccess : IGratuityDataAccess
    {
        private readonly string _connectionString;

        public GratuityDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<GratuitySheetModel> GetGratuitySheetAsync(int year, string? deptId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            // Check if year is posted
            const string checkPostedSql = @"
                SELECT TOP 1 ISNULL(Posted, 0)
                FROM Employees_Gratuity
                WHERE GratuityYear = @Year";

            bool isPosted = await db.ExecuteScalarAsync<bool?>(checkPostedSql, new { Year = year }) ?? false;

            var dtJoin = new DateTime(year, 7, 1);
            var dtDiscontinued = new DateTime(year, 1, 1);

            const string querySql = @"
                SELECT 
                    ISNULL(e.EmpID, '') AS EmpID,
                    LTRIM(RTRIM(ISNULL(e.Name, '') + ' ' + ISNULL(e.Rel, '') + ' ' + ISNULL(e.FName, ''))) AS EmpName,
                    ISNULL(e.Designation, '') AS Designation,
                    ISNULL(e.DeptID, '') AS DeptID,
                    ISNULL(e.DeptName, '') AS DeptName,
                    CASE WHEN ISNULL(ms.BSal, 0) > 0 THEN ISNULL(ms.BSal, 0) ELSE ISNULL(e.StartingSalary, 0) END AS Salary,
                    ISNULL(e.JoinDate, '1900-01-01') AS JoinDate,
                    e.DiscontinuedOn,
                    tg.GratuityAmount AS SavedGratuityAmount
                FROM VEmp e
                LEFT OUTER JOIN (
                    SELECT EmpID, BSal 
                    FROM MonthlySalaries 
                    WHERE EntryID IN (SELECT MAX(EntryID) FROM MonthlySalaries GROUP BY EmpID)
                ) ms ON e.EmpID = ms.EmpID
                LEFT OUTER JOIN (
                    SELECT d.EmpID, d.GratuityAmount
                    FROM Employees_Gratuity_Detail d
                    INNER JOIN Employees_Gratuity g ON g.EntryID = d.RefID
                    WHERE g.GratuityYear = @Year
                ) tg ON e.EmpID = tg.EmpID
                WHERE e.JoinDate < @DTJoin 
                  AND (e.DiscontinuedOn IS NULL OR e.DiscontinuedOn > @DTDiscontinued)
                  AND (@DeptID IS NULL OR @DeptID = '0' OR e.DeptID = @DeptID)
                ORDER BY e.EmpID";

            var rawRows = await db.QueryAsync<GratuityRowModel>(querySql, new
            {
                Year = year,
                DTJoin = dtJoin,
                DTDiscontinued = dtDiscontinued,
                DeptID = string.IsNullOrWhiteSpace(deptId) ? null : deptId
            });

            var sheet = new GratuitySheetModel
            {
                Year = year,
                IsPosted = isPosted,
                Rows = new List<GratuityRowModel>()
            };

            foreach (var r in rawRows)
            {
                int days;
                int months = CalculateServiceMonths(r.JoinDate, r.DiscontinuedOn, year, out days);
                r.MonthsWorked = months;
                r.DaysWorked = days;

                if (r.SavedGratuityAmount.HasValue && r.SavedGratuityAmount.Value > 0)
                {
                    r.GratuityAmount = r.SavedGratuityAmount.Value;
                }
                else
                {
                    decimal rate = r.Salary / 12.0m;
                    r.GratuityAmount = (int)Math.Round(rate * months, MidpointRounding.AwayFromZero);
                }

                sheet.Rows.Add(r);
            }

            return sheet;
        }

        public async Task<bool> SaveGratuitySheetAsync(SaveGratuityDto dto)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            if (db.State != ConnectionState.Open)
            {
                db.Open();
            }

            // Check if year is already posted
            const string checkPostedSql = @"
                SELECT COUNT(1) 
                FROM Employees_Gratuity 
                WHERE GratuityYear = @Year AND ISNULL(Posted, 0) = 1";

            int postedCount = await db.ExecuteScalarAsync<int>(checkPostedSql, new { dto.Year });
            if (postedCount > 0)
            {
                throw new InvalidOperationException("Vouchers have been posted for this year. Editing is Locked.");
            }

            using var trans = db.BeginTransaction();
            try
            {
                // Delete existing details and header for this year
                const string deleteSql = @"
                    DELETE FROM Employees_Gratuity_Detail 
                    WHERE RefID IN (SELECT EntryID FROM Employees_Gratuity WHERE GratuityYear = @Year);

                    DELETE FROM Employees_Gratuity WHERE GratuityYear = @Year;";

                await db.ExecuteAsync(deleteSql, new { dto.Year }, trans);

                // Insert new header
                const string insertHeaderSql = @"
                    INSERT INTO Employees_Gratuity (GratuityYear, Posted) 
                    VALUES (@Year, 0);
                    SELECT SCOPE_IDENTITY();";

                int refId = await db.ExecuteScalarAsync<int>(insertHeaderSql, new { dto.Year }, trans);

                // Insert details
                const string insertDetailSql = @"
                    INSERT INTO Employees_Gratuity_Detail (RefID, EmpID, GratuityAmount)
                    VALUES (@RefID, @EmpID, @GratuityAmount);";

                foreach (var row in dto.Rows)
                {
                    await db.ExecuteAsync(insertDetailSql, new
                    {
                        RefID = refId,
                        EmpID = row.EmpID,
                        GratuityAmount = row.GratuityAmount
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

        public static int CalculateServiceMonths(DateTime joinDate, DateTime? discontinuedOn, int year, out int days)
        {
            DateTime myStartDT = (joinDate <= new DateTime(year, 1, 31))
                ? new DateTime(year, 1, 1)
                : joinDate;

            DateTime myEndDT = (!discontinuedOn.HasValue || discontinuedOn.Value >= new DateTime(year, 12, 31))
                ? new DateTime(year, 12, 31)
                : discontinuedOn.Value;

            if (myStartDT > myEndDT)
            {
                days = 0;
                return 0;
            }

            DateTime endDT = myEndDT.AddDays(1);
            int iYears = endDT.Year - myStartDT.Year;
            int iMonths = endDT.Month - myStartDT.Month + (iYears * 12);

            if (myStartDT.Day > endDT.Day)
            {
                iMonths--;
            }

            iYears = iMonths / 12;
            days = (int)(endDT - myStartDT.AddMonths(iMonths)).TotalDays;
            if (days < 0)
            {
                days = 0;
            }

            return Math.Max(0, iMonths);
        }
    }
}
