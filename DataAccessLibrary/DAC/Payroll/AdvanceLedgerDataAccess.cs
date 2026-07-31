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
    }
}
