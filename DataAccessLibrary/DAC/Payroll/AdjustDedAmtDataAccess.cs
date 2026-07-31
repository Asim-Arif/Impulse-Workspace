using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class AdjustDedAmtDataAccess : IAdjustDedAmtDataAccess
    {
        private readonly string _connectionString;

        public AdjustDedAmtDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<EmpDeductionInfoModel> GetEmployeeLatestLoanDeductionAsync(string empId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                SELECT TOP 1
                    a.EntryID,
                    a.EmpID,
                    ISNULL(a.DAmount, 0) AS CurrentDAmount,
                    ISNULL(a.Amount, 0) AS LoanAmount,
                    ISNULL(a.Amount, 0) - ISNULL((SELECT SUM(AmtClrd) FROM AmtCleared WHERE EmpID = a.EmpID), 0) AS Balance,
                    a.DT AS IssueDate,
                    a.Description
                FROM Advances a
                WHERE a.Type = 1 AND a.EmpID = @EmpID
                ORDER BY a.EntryID DESC";

            var result = await db.QueryFirstOrDefaultAsync<EmpDeductionInfoModel>(sql, new { EmpID = empId });

            if (result != null)
            {
                result.HasActiveLoan = true;
                return result;
            }

            return new EmpDeductionInfoModel
            {
                EmpID = empId,
                HasActiveLoan = false
            };
        }

        public async Task UpdateDeductionAmountAsync(long entryId, decimal newDAmount)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                UPDATE Advances
                SET DAmount = @NewDAmount
                WHERE EntryID = @EntryID";

            await db.ExecuteAsync(sql, new { NewDAmount = newDAmount, EntryID = entryId });
        }
    }
}
