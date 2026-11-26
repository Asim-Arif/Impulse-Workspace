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
    public class AdvanceLedgerDataAccess : IAdvanceLedgerDataAccess
    {
        private readonly string _connectionString;

        public AdvanceLedgerDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<AdvanceLedgerItemModel>> GetAdvanceLedgerAsync(AdvanceLedgerFilterDto filter)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string sql = @"
                SELECT
                    v.EntryID,
                    v.EmpID,
                    v.Name AS EmpName,
                    v.DeptID,
                    v.DeptName,
                    v.Designation,
                    v.DT,
                    v.Type,
                    ISNULL(v.Amount, 0) AS Amount,
                    ISNULL(v.DAmount, 0) AS DAmount,
                    ISNULL(v.AmountCleared, 0) AS AmountCleared,
                    ISNULL(v.Amount, 0) - ISNULL(v.AmountCleared, 0) AS Balance,
                    v.AccVoucherNo,
                    ISNULL(v.Cleared, 0) AS Cleared,
                    v.DeductionStartDT
                FROM VAdvances v
                WHERE 1 = 1";

            var parameters = new DynamicParameters();

            if (filter.LoanType >= 0)
            {
                sql += " AND v.Type = @LoanType";
                parameters.Add("LoanType", filter.LoanType);
            }

            if (!string.IsNullOrEmpty(filter.DeptID) && filter.DeptID != "0")
            {
                sql += " AND v.DeptID = @DeptID";
                parameters.Add("DeptID", filter.DeptID);
            }

            if (!string.IsNullOrEmpty(filter.EmpID) && filter.EmpID != "0")
            {
                sql += " AND v.EmpID = @EmpID";
                parameters.Add("EmpID", filter.EmpID);
            }

            sql += " AND v.DT BETWEEN @FromDate AND @ToDate";
            parameters.Add("FromDate", filter.FromDate.Date);
            parameters.Add("ToDate", filter.ToDate.Date.AddDays(1).AddTicks(-1));

            sql += " ORDER BY v.EmpID, v.DT, v.Type ASC";

            var result = await db.QueryAsync<AdvanceLedgerItemModel>(sql, parameters);
            return result.ToList();
        }

        public async Task<decimal> GetLongTermOpeningBalanceAsync(string empId, DateTime fromDate)
        {
            if (string.IsNullOrWhiteSpace(empId) || empId == "0")
                return 0m;

            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                SELECT
                    ISNULL((SELECT SUM(CAST(Amount AS DECIMAL(18, 2))) FROM Advances WHERE EmpID = @EmpID AND Type = 1 AND DT < @FromDate), 0)
                  - ISNULL((SELECT SUM(CAST(AmtClrd AS DECIMAL(18, 2))) FROM AmtCleared WHERE EmpID = @EmpID AND DT < @FromDate), 0)";

            return await db.ExecuteScalarAsync<decimal>(sql, new
            {
                EmpID = empId,
                FromDate = fromDate.Date
            });
        }
    }
}
