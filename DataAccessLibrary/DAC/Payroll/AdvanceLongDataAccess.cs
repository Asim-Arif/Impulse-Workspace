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
    public class AdvanceLongDataAccess : IAdvanceLongDataAccess
    {
        private readonly string _connectionString;

        public AdvanceLongDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        /// <summary>
        /// Returns outstanding long-term loan balance and last monthly deduction
        /// for the given employee, matching the original VB6 query logic.
        /// </summary>
        public async Task<AdvanceLongSummaryModel> GetEmployeeAdvanceSummaryAsync(string empId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            // Outstanding balance = SUM(Amount) - SUM(AmountCleared from AmtCleared)
            const string balanceSql = @"
                SELECT ISNULL(SUM(a.Amount), 0) -
                       ISNULL((SELECT SUM(AmtClrd) FROM AmtCleared WHERE EmpID = @EmpID), 0)
                FROM Advances a
                WHERE a.EmpID = @EmpID AND a.Type = 1";

            decimal balance = await db.ExecuteScalarAsync<decimal>(balanceSql, new { EmpID = empId });

            // Last deduction amount from the most recent Type=1 advance
            const string dedSql = @"
                SELECT ISNULL(DAmount, 0)
                FROM Advances
                WHERE EntryID = (SELECT MAX(EntryID) FROM Advances WHERE Type = 1 AND EmpID = @EmpID)";

            decimal lastDed = await db.ExecuteScalarAsync<decimal>(dedSql, new { EmpID = empId });

            return new AdvanceLongSummaryModel
            {
                PreviousBalance         = balance,
                PreviousDeductionAmount = lastDed
            };
        }

        /// <summary>
        /// Inserts a row into the Advances table with Type=1 (Long Term).
        /// The voucher has already been posted by the service layer and
        /// its VchrNo is passed in here to be stored in AccVoucherNo.
        /// </summary>
        public async Task SaveAdvanceLongRecordAsync(AdvanceLongDto dto, string postedVchrNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                INSERT INTO Advances
                    (EmpID, DT, Type, Description, Amount, DAmount, AmountCleared, AccVoucherNo, DeductionStartDT)
                VALUES
                    (@EmpID, @DT, 1, @Description, @Amount, @DAmount, 0, @AccVoucherNo, @DeductionStartDT)";

            await db.ExecuteAsync(sql, new
            {
                EmpID            = dto.EmpID,
                DT               = dto.IssueDate,
                Description      = dto.Description,
                Amount           = dto.Amount,
                DAmount          = dto.DAmount,
                AccVoucherNo     = postedVchrNo,
                DeductionStartDT = dto.DeductionStartDT
            });
        }
    }
}
