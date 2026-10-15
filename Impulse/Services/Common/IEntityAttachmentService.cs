using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Attachments;
using Microsoft.AspNetCore.Components.Forms;

namespace Impulse.Services.Common
{
    public interface IEntityAttachmentService
    {
        Task<List<EntityAttachmentDto>> GetAttachmentsAsync(string entityType, string entityId, string entitySecondaryId = null);
        Task<EntityAttachmentDto> UploadAttachmentAsync(string entityType, string entityId, string entitySecondaryId, IBrowserFile file, string remarks = null, string uploadedBy = null, long maxFileSize = 52428800); // 50MB default
        Task<bool> DeleteAttachmentAsync(long id, string deletedBy = null);
        Task<bool> UpdateRemarksAsync(long id, string remarks);
        Task<int> GetCountAsync(string entityType, string entityId, string entitySecondaryId = null);
        string GetPhysicalFilePath(EntityAttachmentDto attachment);
    }
}
