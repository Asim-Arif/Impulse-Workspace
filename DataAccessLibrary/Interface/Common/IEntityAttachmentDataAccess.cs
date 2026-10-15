using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Attachments;

namespace DataAccessLibrary.Interface.Common
{
    public interface IEntityAttachmentDataAccess
    {
        Task<List<EntityAttachmentDto>> GetAttachmentsAsync(string entityType, string entityId, string? secondaryId = null);
        Task<EntityAttachmentDto?> GetAttachmentByIdAsync(long id);
        Task<long> InsertAttachmentAsync(EntityAttachmentDto attachment);
        Task<bool> SoftDeleteAttachmentAsync(long id, string deletedBy);
        Task<int> GetAttachmentCountAsync(string entityType, string entityId, string? secondaryId = null);
    }
}
