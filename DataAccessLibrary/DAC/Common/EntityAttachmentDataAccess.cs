using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Models.ViewModels.Attachments;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Common
{
    public class EntityAttachmentDataAccess : IEntityAttachmentDataAccess
    {
        private readonly IConfiguration _configuration;

        public EntityAttachmentDataAccess(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string ConnectionString => _configuration.GetConnectionString("DefaultConnection")
            ?? _configuration.GetConnectionString("ImpulseDatabase")
            ?? _configuration.GetConnectionString("ImpulseConnection")
            ?? string.Empty;

        public async Task<List<EntityAttachmentDto>> GetAttachmentsAsync(string entityType, string entityId, string? secondaryId = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT Id, EntityType, EntityId, EntitySecondaryId, OriginalFileName, StoredFileName,
                       FilePath, FileSize, FileExtension, ContentType, Remarks, UploadedBy, UploadedAt, IsDeleted
                FROM EntityAttachments
                WHERE EntityType = @EntityType 
                  AND EntityId = @EntityId 
                  AND (EntitySecondaryId = @SecondaryId OR (@SecondaryId IS NULL AND EntitySecondaryId IS NULL))
                  AND IsDeleted = 0
                ORDER BY UploadedAt DESC";

            var result = await db.QueryAsync<EntityAttachmentDto>(sql, new { EntityType = entityType, EntityId = entityId, SecondaryId = secondaryId });
            return result.ToList();
        }

        public async Task<EntityAttachmentDto?> GetAttachmentByIdAsync(long id)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT Id, EntityType, EntityId, EntitySecondaryId, OriginalFileName, StoredFileName,
                       FilePath, FileSize, FileExtension, ContentType, Remarks, UploadedBy, UploadedAt, IsDeleted
                FROM EntityAttachments
                WHERE Id = @Id AND IsDeleted = 0";

            return await db.QueryFirstOrDefaultAsync<EntityAttachmentDto>(sql, new { Id = id });
        }

        public async Task<long> InsertAttachmentAsync(EntityAttachmentDto attachment)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                INSERT INTO EntityAttachments (
                    EntityType, EntityId, EntitySecondaryId, OriginalFileName, StoredFileName,
                    FilePath, FileSize, FileExtension, ContentType, Remarks, UploadedBy, UploadedAt, IsDeleted
                ) VALUES (
                    @EntityType, @EntityId, @EntitySecondaryId, @OriginalFileName, @StoredFileName,
                    @FilePath, @FileSize, @FileExtension, @ContentType, @Remarks, @UploadedBy, GETDATE(), 0
                );
                SELECT CAST(SCOPE_IDENTITY() AS BIGINT);";

            return await db.ExecuteScalarAsync<long>(sql, attachment);
        }

        public async Task<bool> SoftDeleteAttachmentAsync(long id, string deletedBy)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                UPDATE EntityAttachments
                SET IsDeleted = 1, DeletedBy = @DeletedBy, DeletedAt = GETDATE()
                WHERE Id = @Id";

            var rows = await db.ExecuteAsync(sql, new { Id = id, DeletedBy = deletedBy });
            return rows > 0;
        }

        public async Task<int> GetAttachmentCountAsync(string entityType, string entityId, string? secondaryId = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT COUNT(1)
                FROM EntityAttachments
                WHERE EntityType = @EntityType 
                  AND EntityId = @EntityId 
                  AND (EntitySecondaryId = @SecondaryId OR (@SecondaryId IS NULL AND EntitySecondaryId IS NULL))
                  AND IsDeleted = 0";

            return await db.ExecuteScalarAsync<int>(sql, new { EntityType = entityType, EntityId = entityId, SecondaryId = secondaryId });
        }
    }
}
