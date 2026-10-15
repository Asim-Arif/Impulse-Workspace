using System;
using System.Data;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Models.ViewModels.Attachments;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Common
{
    public class EntityNoteDataAccess : IEntityNoteDataAccess
    {
        private readonly IConfiguration _configuration;

        public EntityNoteDataAccess(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string ConnectionString => _configuration.GetConnectionString("DefaultConnection")
            ?? _configuration.GetConnectionString("ImpulseDatabase")
            ?? _configuration.GetConnectionString("ImpulseConnection")
            ?? string.Empty;

        public async Task<EntityNoteDto?> GetNoteAsync(string entityType, string entityId, string? secondaryId = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT Id, EntityType, EntityId, EntitySecondaryId, NoteText, LastModifiedBy, LastModifiedAt
                FROM EntityNotes
                WHERE EntityType = @EntityType 
                  AND EntityId = @EntityId 
                  AND (EntitySecondaryId = @SecondaryId OR (@SecondaryId IS NULL AND EntitySecondaryId IS NULL))";

            return await db.QueryFirstOrDefaultAsync<EntityNoteDto>(sql, new { EntityType = entityType, EntityId = entityId, SecondaryId = secondaryId });
        }

        public async Task<bool> SaveNoteAsync(string entityType, string entityId, string? secondaryId, string noteText, string user)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                IF EXISTS (
                    SELECT 1 FROM EntityNotes 
                    WHERE EntityType = @EntityType 
                      AND EntityId = @EntityId 
                      AND (EntitySecondaryId = @SecondaryId OR (@SecondaryId IS NULL AND EntitySecondaryId IS NULL))
                )
                BEGIN
                    UPDATE EntityNotes
                    SET NoteText = @NoteText, LastModifiedBy = @User, LastModifiedAt = GETDATE()
                    WHERE EntityType = @EntityType 
                      AND EntityId = @EntityId 
                      AND (EntitySecondaryId = @SecondaryId OR (@SecondaryId IS NULL AND EntitySecondaryId IS NULL));
                END
                ELSE
                BEGIN
                    INSERT INTO EntityNotes (EntityType, EntityId, EntitySecondaryId, NoteText, LastModifiedBy, LastModifiedAt)
                    VALUES (@EntityType, @EntityId, @SecondaryId, @NoteText, @User, GETDATE());
                END";

            var rows = await db.ExecuteAsync(sql, new
            {
                EntityType = entityType,
                EntityId = entityId,
                SecondaryId = secondaryId,
                NoteText = noteText,
                User = user
            });
            return rows > 0;
        }
    }
}
