using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class ClearShortTermDataAccess : IClearShortTermDataAccess
    {
        private readonly string _connectionString;

        public ClearShortTermDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        /// <summary>
        /// Open short-term advances for the employee, matching the legacy
        /// frmClearShortTerm query (Type=0, Cleared=0, ordered by EntryID).
        /// </summary>
        public async Task<List<ShortTermAdvanceItem>> GetOpenAdvancesAsync(string empId)
        {
            using var db = new SqlConnection(_connectionString);

            const string sql = @"
                SELECT EntryID, DT AS Date, Amount, ISNULL(Description, '') AS Description
                FROM Advances
                WHERE EmpID = @EmpID AND Type = 0 AND Cleared = 0
                ORDER BY EntryID";

            var rows = await db.QueryAsync<ShortTermAdvanceItem>(sql, new { EmpID = empId });
            return rows.ToList();
        }

        /// <summary>
        /// Inserts the clearance header and, per selected advance, sets
        /// Cleared=2 and writes the detail row — all in a single transaction
        /// (as the legacy form did). The detail Amount is populated with the
        /// advance's full amount (the ledger views SUM it; legacy never wrote it).
        /// </summary>
        public async Task SaveClearanceAsync(ClearShortTermDto dto, string vchrNo, string userName, string machineName)
        {
            using var db = new SqlConnection(_connectionString);
            db.Open();
            using var tx = db.BeginTransaction();

            const string insertHeaderSql = @"
                INSERT INTO ShortTermAdvanceClearance
                    (EmpID, DT, Remarks, UserName, MachineName, Type, VchrNo)
                VALUES
                    (@EmpID, @DT, @Remarks, @UserName, @MachineName, 0, @VchrNo);
                SELECT CAST(SCOPE_IDENTITY() AS int);";

            int sacId = await db.ExecuteScalarAsync<int>(insertHeaderSql, new
            {
                EmpID       = dto.EmpID,
                DT          = dto.ClearDate,
                Remarks     = dto.Remarks,
                UserName    = userName,
                MachineName = machineName,
                VchrNo      = vchrNo
            }, transaction: tx);

            const string updateAdvanceSql = @"
                UPDATE Advances
                SET Cleared = 2
                WHERE EntryID = @EntryID";

            const string insertDetailSql = @"
                INSERT INTO ShortTermAdvanceClearanceDetail
                    (SAC_RefID, Advances_RefID, Amount)
                VALUES
                    (@SAC_RefID, @Advances_RefID, @Amount)";

            foreach (var adv in dto.SelectedAdvances)
            {
                await db.ExecuteAsync(updateAdvanceSql, new { EntryID = adv.EntryID }, transaction: tx);

                await db.ExecuteAsync(insertDetailSql, new
                {
                    SAC_RefID       = sacId,
                    Advances_RefID  = adv.EntryID,
                    Amount          = (int)adv.Amount
                }, transaction: tx);
            }

            tx.Commit();
        }
    }
}
