using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class ClearAdvanceDataAccess : IClearAdvanceDataAccess
    {
        private readonly string _connectionString;

        public ClearAdvanceDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        /// <summary>
        /// Returns outstanding long-term loan balance, current monthly deduction
        /// and latest advance EntryID for the given employee,
        /// matching the original VB6 query logic (frmClearAmt).
        /// </summary>
        public async Task<ClearAdvanceSummaryModel> GetSummaryAsync(string empId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            // Outstanding balance = SUM(Advances.Amount, Type=1) - SUM(AmtCleared.AmtClrd)
            // (AmtCleared has no Type filter in the legacy query)
            const string balanceSql = @"
                SELECT ISNULL((SELECT SUM(Amount) FROM Advances WHERE EmpID = @EmpID AND Type = 1), 0) -
                       ISNULL((SELECT SUM(AmtClrd) FROM AmtCleared WHERE EmpID = @EmpID), 0)";

            decimal balance = await db.ExecuteScalarAsync<decimal>(balanceSql, new { EmpID = empId });

            // Latest Type=1 advance (may not exist for this employee)
            const string latestSql = @"
                SELECT EntryID, ISNULL(DAmount, 0) AS DAmount
                FROM Advances
                WHERE EntryID = (SELECT MAX(EntryID) FROM Advances WHERE Type = 1 AND EmpID = @EmpID)";

            var latest = await db.QueryFirstOrDefaultAsync<LatestAdvanceRow>(latestSql, new { EmpID = empId });

            return new ClearAdvanceSummaryModel
            {
                OutstandingBalance   = balance,
                CurrentDeduction     = latest?.DAmount ?? 0,
                LatestAdvanceEntryID = latest?.EntryID
            };
        }

        /// <summary>
        /// Inserts the AmtCleared row and conditionally updates the latest
        /// advance's DAmount, in a single transaction (as the legacy form did).
        /// </summary>
        public async Task SaveClearanceAsync(ClearAdvanceDto dto, string vchrNo)
        {
            using var db = new SqlConnection(_connectionString);
            db.Open();
            using var tx = db.BeginTransaction();

            const string insertSql = @"
                INSERT INTO AmtCleared
                    (EmpID, DT, AmtClrd, VchrNo, Description)
                VALUES
                    (@EmpID, @DT, @AmtClrd, @VchrNo, @Description)";

            await db.ExecuteAsync(insertSql, new
            {
                EmpID       = dto.EmpID,
                DT          = dto.ClearDate,
                AmtClrd     = dto.AmountReceived,
                VchrNo      = vchrNo,
                Description = dto.Description
            }, transaction: tx);

            if (dto.NewDeduction > 0 && dto.LatestAdvanceEntryID.HasValue)
            {
                const string updateSql = @"
                    UPDATE Advances
                    SET DAmount = @NewDeduction
                    WHERE EntryID = @EntryID";

                await db.ExecuteAsync(updateSql, new
                {
                    NewDeduction = dto.NewDeduction,
                    EntryID      = dto.LatestAdvanceEntryID.Value
                }, transaction: tx);
            }

            tx.Commit();
        }

        private class LatestAdvanceRow
        {
            public long    EntryID { get; set; }
            public decimal DAmount { get; set; }
        }
    }
}
