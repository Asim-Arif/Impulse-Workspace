using System;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Models.ViewModels.Attachments;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Common
{
    public class EntityNoteService : IEntityNoteService
    {
        private readonly IEntityNoteDataAccess _noteDac;
        private readonly ILogger<EntityNoteService> _logger;

        public EntityNoteService(
            IEntityNoteDataAccess noteDac,
            ILogger<EntityNoteService> logger)
        {
            _noteDac = noteDac;
            _logger = logger;
        }

        public async Task<EntityNoteDto?> GetNoteAsync(string entityType, string entityId, string entitySecondaryId = null)
        {
            try
            {
                return await _noteDac.GetNoteAsync(entityType, entityId, entitySecondaryId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting note for {EntityType} {EntityId}", entityType, entityId);
                return null;
            }
        }

        public async Task<bool> SaveNoteAsync(string entityType, string entityId, string entitySecondaryId, string noteText, string user = null)
        {
            try
            {
                return await _noteDac.SaveNoteAsync(entityType, entityId, entitySecondaryId, noteText ?? string.Empty, user ?? "System");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving note for {EntityType} {EntityId}", entityType, entityId);
                return false;
            }
        }
    }
}
