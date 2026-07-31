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
    public class AdvanceShortDataAccess : IAdvanceShortDataAccess
    {
        private readonly string _connectionString;

        public AdvanceShortDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        /// <summary>
        /// Returns outstanding short-term advance balance (Type = 0) for the given employee.
        /// </summary>
        public async Task<AdvanceShortSummaryModel> GetEmployeeAdvanceSummaryAsync(string empId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            // Outstanding short term balance = SUM(Amount) - SUM(AmountCleared) for Type = 0
            const string balanceSql = @"
                SELECT ISNULL(SUM(a.Amount - ISNULL(a.AmountCleared, 0)), 0)
                FROM Advances a
                WHERE a.EmpID = @EmpID AND a.Type = 0 AND (a.Cleared IS NULL OR a.Cleared = 0)";

            decimal balance = await db.ExecuteScalarAsync<decimal>(balanceSql, new { EmpID = empId });

            return new AdvanceShortSummaryModel
            {
                PreviousShortTermBalance = balance
            };
        }

        /// <summary>
        /// Inserts a row into the Advances table with Type=0 (Short Term) and DAmount=0.
        /// The voucher has already been posted by the service layer and
        /// its VchrNo is passed in here to be stored in AccVoucherNo.
        /// </summary>
        public async Task SaveAdvanceShortRecordAsync(AdvanceShortDto dto, string postedVchrNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                INSERT INTO Advances
                    (EmpID, DT, Type, Description, Amount, DAmount, AmountCleared, AccVoucherNo, DeductionStartDT, ActualAmountTaken)
                VALUES
                    (@EmpID, @DT, 0, @Description, @Amount, 0, 0, @AccVoucherNo, @DeductionStartDT, @Amount)";

            await db.ExecuteAsync(sql, new
            {
                EmpID            = dto.EmpID,
                DT               = dto.IssueDate,
                Description      = dto.Description,
                Amount           = dto.Amount,
                AccVoucherNo     = postedVchrNo,
                DeductionStartDT = dto.DeductionStartDT
            });
        }
    }
}
