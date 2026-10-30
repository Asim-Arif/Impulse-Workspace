using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.QMS
{
    public class QMSReasonDataAccess : IQMSReasonDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<QMSReasonDataAccess> _logger;

        public QMSReasonDataAccess(IConfiguration configuration, ILogger<QMSReasonDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<QMSReasonModel>> GetAllReasonsAsync(string? searchText = null)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        r.EntryID,
                        r.Reason_QMS,
                        (SELECT COUNT(1) FROM CustomerComplaints_Detail ccd WHERE ccd.ItemReason_Main = r.Reason_QMS) AS UsageCount
                    FROM QMS_Reasons_List r
                    WHERE (@Search IS NULL OR r.Reason_QMS LIKE '%' + @Search + '%')
                    ORDER BY r.Reason_QMS;";

                var param = string.IsNullOrWhiteSpace(searchText) ? null : searchText.Trim();
                var results = await db.QueryAsync<QMSReasonModel>(sql, new { Search = param });
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS Reasons");
                throw;
            }
        }

        public async Task<QMSReasonModel?> GetReasonByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        r.EntryID,
                        r.Reason_QMS,
                        (SELECT COUNT(1) FROM CustomerComplaints_Detail ccd WHERE ccd.ItemReason_Main = r.Reason_QMS) AS UsageCount
                    FROM QMS_Reasons_List r
                    WHERE r.EntryID = @EntryID;";

                return await db.QueryFirstOrDefaultAsync<QMSReasonModel>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS Reason ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<bool> ReasonExistsAsync(string reason, int? excludeEntryId = null)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT COUNT(1)
                    FROM QMS_Reasons_List
                    WHERE LTRIM(RTRIM(LOWER(Reason_QMS))) = LTRIM(RTRIM(LOWER(@Reason)))
                      AND (@ExcludeId IS NULL OR EntryID <> @ExcludeId);";

                int count = await db.ExecuteScalarAsync<int>(sql, new
                {
                    Reason = reason.Trim(),
                    ExcludeId = excludeEntryId
                });

                return count > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking if QMS Reason exists: {Reason}", reason);
                throw;
            }
        }

        public async Task<int> AddReasonAsync(string reason)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    INSERT INTO QMS_Reasons_List (Reason_QMS)
                    VALUES (@Reason);
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                return await db.ExecuteScalarAsync<int>(sql, new { Reason = reason.Trim() });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding QMS Reason: {Reason}", reason);
                throw;
            }
        }

        public async Task<bool> UpdateReasonAsync(int entryId, string originalReason, string newReason)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var tx = db.BeginTransaction();

            try
            {
                var cleanNewReason = newReason.Trim();
                var cleanOldReason = originalReason.Trim();

                if (!string.Equals(cleanOldReason, cleanNewReason, StringComparison.OrdinalIgnoreCase))
                {
                    // Cascade update references in CustomerComplaints_Detail
                    const string cascadeSql = @"
                        UPDATE CustomerComplaints_Detail 
                        SET ItemReason_Main = @NewReason 
                        WHERE ItemReason_Main = @OldReason;";
                    await db.ExecuteAsync(cascadeSql, new { NewReason = cleanNewReason, OldReason = cleanOldReason }, tx);
                }

                const string updateSql = @"
                    UPDATE QMS_Reasons_List 
                    SET Reason_QMS = @NewReason 
                    WHERE (EntryID = @EntryID AND @EntryID > 0) OR Reason_QMS = @OldReason;";

                int affected = await db.ExecuteAsync(updateSql, new
                {
                    NewReason = cleanNewReason,
                    OldReason = cleanOldReason,
                    EntryID = entryId
                }, tx);

                tx.Commit();
                return affected > 0;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error updating QMS Reason ID {EntryID} from '{OldReason}' to '{NewReason}'", entryId, originalReason, newReason);
                throw;
            }
        }

        public async Task<bool> DeleteReasonAsync(int entryId, string reason)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    DELETE FROM QMS_Reasons_List 
                    WHERE (EntryID = @EntryID AND @EntryID > 0) 
                       OR Reason_QMS = @Reason;";

                int affected = await db.ExecuteAsync(sql, new { EntryID = entryId, Reason = reason.Trim() });
                return affected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting QMS Reason: {Reason} (ID {EntryID})", reason, entryId);
                throw;
            }
        }

        public async Task<int> GetUsageCountAsync(string reason)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT COUNT(1)
                    FROM CustomerComplaints_Detail
                    WHERE ItemReason_Main = @Reason;";

                return await db.ExecuteScalarAsync<int>(sql, new { Reason = reason.Trim() });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting usage count for Reason: {Reason}", reason);
                return 0;
            }
        }
    }
}
