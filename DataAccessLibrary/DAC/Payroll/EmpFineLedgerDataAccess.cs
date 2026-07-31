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
    public class EmpFineLedgerDataAccess : IEmpFineLedgerDataAccess
    {
        private readonly string _connectionString;

        public EmpFineLedgerDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<decimal> GetOpeningBalanceAsync(string empId, DateTime fromDate)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string fineSql;
            string clrSql;

            if (!string.IsNullOrEmpty(empId) && empId != "0")
            {
                clrSql = "(SELECT ISNULL(SUM(AmtClrd), 0) FROM EmpFineAmtCleared WHERE DT < @FromDate AND EmpID = @EmpID)";
                fineSql = $"SELECT ISNULL(SUM(Amount), 0) - {clrSql} FROM EmpFine WHERE FineDate < @FromDate AND EmpID = @EmpID";
            }
            else
            {
                clrSql = "(SELECT ISNULL(SUM(AmtClrd), 0) FROM EmpFineAmtCleared WHERE DT < @FromDate)";
                fineSql = $"SELECT ISNULL(SUM(Amount), 0) - {clrSql} FROM EmpFine WHERE FineDate < @FromDate";
            }

            return await db.ExecuteScalarAsync<decimal>(fineSql, new { EmpID = empId, FromDate = fromDate.Date });
        }

        public async Task<List<EmpFineLedgerItemModel>> GetFineLedgerAsync(EmpFineLedgerFilterDto filter)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string sql = @"
                SELECT
                    v.EntryID,
                    v.empid AS EmpID,
                    v.name AS EmpName,
                    v.deptid AS DeptID,
                    v.DeptName,
                    v.Designation,
                    v.DT,
                    ISNULL(v.Description, '') AS Description,
                    ISNULL(v.Amount, 0) AS Amount
                FROM VEmpFineLedger v
                WHERE v.DT BETWEEN @FromDate AND @ToDate";

            var parameters = new DynamicParameters();

            if (!string.IsNullOrEmpty(filter.DeptID) && filter.DeptID != "0")
            {
                sql += " AND v.deptid = @DeptID";
                parameters.Add("DeptID", filter.DeptID);
            }

            if (!string.IsNullOrEmpty(filter.EmpID) && filter.EmpID != "0")
            {
                sql += " AND v.empid = @EmpID";
                parameters.Add("EmpID", filter.EmpID);
            }

            sql += " ORDER BY v.empid, v.DT ASC";

            parameters.Add("FromDate", filter.FromDate.Date);
            parameters.Add("ToDate", filter.ToDate.Date.AddDays(1).AddTicks(-1));

            var result = await db.QueryAsync<EmpFineLedgerItemModel>(sql, parameters);
            return result.ToList();
        }

        public async Task DeleteFineAsync(long entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            // Fetch FineDate for the target entry
            var fineDate = await db.ExecuteScalarAsync<DateTime?>("SELECT FineDate FROM EmpFine WHERE FineID = @EntryID", new { EntryID = entryId });
            if (fineDate.HasValue)
            {
                // Check if salary for that month is finalized
                DateTime monthEnd = new DateTime(fineDate.Value.Year, fineDate.Value.Month, DateTime.DaysInMonth(fineDate.Value.Year, fineDate.Value.Month));
                int salCount = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM MonthlySalaries WHERE DT = @MonthEnd", new { MonthEnd = monthEnd });

                if (salCount > 0)
                {
                    throw new InvalidOperationException("Cannot delete fine record because salary for this month has been finalized.");
                }
            }

            await db.ExecuteAsync("DELETE FROM EmpFine WHERE FineID = @EntryID", new { EntryID = entryId });
        }
    }
}
