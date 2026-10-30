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
    public class QMSDocumentTypeDataAccess : IQMSDocumentTypeDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<QMSDocumentTypeDataAccess> _logger;

        public QMSDocumentTypeDataAccess(IConfiguration configuration, ILogger<QMSDocumentTypeDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<QMSDocumentTypeModel>> GetDocumentTypesAsync(string? searchTerm = null)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        dt.EntryID,
                        dt.DocumentType,
                        (SELECT COUNT(1) FROM QMS_Documents d WHERE d.DocumentType = dt.DocumentType) AS UsageCount
                    FROM QMS_DocumentTypes dt
                    WHERE (@Search IS NULL OR dt.DocumentType LIKE '%' + @Search + '%')
                    ORDER BY dt.DocumentType;";

                var param = string.IsNullOrWhiteSpace(searchTerm) ? null : searchTerm.Trim();
                var results = await db.QueryAsync<QMSDocumentTypeModel>(sql, new { Search = param });
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS Document Types");
                throw;
            }
        }

        public async Task<QMSDocumentTypeModel?> GetDocumentTypeByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        dt.EntryID,
                        dt.DocumentType,
                        (SELECT COUNT(1) FROM QMS_Documents d WHERE d.DocumentType = dt.DocumentType) AS UsageCount
                    FROM QMS_DocumentTypes dt
                    WHERE dt.EntryID = @EntryID;";

                return await db.QueryFirstOrDefaultAsync<QMSDocumentTypeModel>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS Document Type ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<bool> DocumentTypeExistsAsync(string documentType, int? excludeEntryId = null)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT COUNT(1)
                    FROM QMS_DocumentTypes
                    WHERE LTRIM(RTRIM(LOWER(DocumentType))) = LTRIM(RTRIM(LOWER(@DocumentType)))
                      AND (@ExcludeId IS NULL OR EntryID <> @ExcludeId);";

                int count = await db.ExecuteScalarAsync<int>(sql, new
                {
                    DocumentType = documentType.Trim(),
                    ExcludeId = excludeEntryId
                });

                return count > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking if QMS Document Type exists: {DocumentType}", documentType);
                throw;
            }
        }

        public async Task<int> AddDocumentTypeAsync(string documentType)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    INSERT INTO QMS_DocumentTypes (DocumentType)
                    VALUES (@DocumentType);
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                return await db.ExecuteScalarAsync<int>(sql, new { DocumentType = documentType.Trim() });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding QMS Document Type: {DocumentType}", documentType);
                throw;
            }
        }

        public async Task<bool> UpdateDocumentTypeAsync(int entryId, string newDocumentType)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var tx = db.BeginTransaction();

            try
            {
                // Get previous name to cascade update into QMS_Documents if needed
                const string getOldSql = "SELECT DocumentType FROM QMS_DocumentTypes WHERE EntryID = @EntryID;";
                var oldName = await db.ExecuteScalarAsync<string>(getOldSql, new { EntryID = entryId }, tx);

                if (!string.IsNullOrWhiteSpace(oldName) && !oldName.Equals(newDocumentType.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    const string cascadeSql = "UPDATE QMS_Documents SET DocumentType = @NewName WHERE DocumentType = @OldName;";
                    await db.ExecuteAsync(cascadeSql, new { NewName = newDocumentType.Trim(), OldName = oldName }, tx);
                }

                const string updateSql = "UPDATE QMS_DocumentTypes SET DocumentType = @DocumentType WHERE EntryID = @EntryID;";
                int affected = await db.ExecuteAsync(updateSql, new
                {
                    DocumentType = newDocumentType.Trim(),
                    EntryID = entryId
                }, tx);

                tx.Commit();
                return affected > 0;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error updating QMS Document Type ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<bool> DeleteDocumentTypeAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "DELETE FROM QMS_DocumentTypes WHERE EntryID = @EntryID;";
                int affected = await db.ExecuteAsync(sql, new { EntryID = entryId });
                return affected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting QMS Document Type ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<int> GetUsageCountAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT COUNT(1)
                    FROM QMS_Documents d
                    INNER JOIN QMS_DocumentTypes dt ON d.DocumentType = dt.DocumentType
                    WHERE dt.EntryID = @EntryID;";

                return await db.ExecuteScalarAsync<int>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting usage count for Document Type ID {EntryID}", entryId);
                return 0;
            }
        }
    }
}

