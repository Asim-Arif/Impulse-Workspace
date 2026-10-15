using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Models.ViewModels.Attachments;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Common
{
    public class EntityAttachmentService : IEntityAttachmentService
    {
        private readonly IEntityAttachmentDataAccess _attachmentDac;
        private readonly IWebHostEnvironment _env;
        private readonly ILogger<EntityAttachmentService> _logger;

        public EntityAttachmentService(
            IEntityAttachmentDataAccess attachmentDac,
            IWebHostEnvironment env,
            ILogger<EntityAttachmentService> logger)
        {
            _attachmentDac = attachmentDac;
            _env = env;
            _logger = logger;
        }

        public async Task<List<EntityAttachmentDto>> GetAttachmentsAsync(string entityType, string entityId, string entitySecondaryId = null)
        {
            try
            {
                return await _attachmentDac.GetAttachmentsAsync(entityType, entityId, entitySecondaryId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting attachments for {EntityType} {EntityId}", entityType, entityId);
                return new List<EntityAttachmentDto>();
            }
        }

        public async Task<EntityAttachmentDto> UploadAttachmentAsync(
            string entityType,
            string entityId,
            string entitySecondaryId,
            IBrowserFile file,
            string remarks = null,
            string uploadedBy = null,
            long maxFileSize = 52428800)
        {
            if (file == null) throw new ArgumentNullException(nameof(file));

            var safeName = Path.GetFileName(file.Name);
            var extension = Path.GetExtension(safeName)?.ToLowerInvariant() ?? "";
            var storedFileName = $"{Guid.NewGuid():N}_{safeName}";

            var webRoot = _env.WebRootPath;
            if (string.IsNullOrWhiteSpace(webRoot))
            {
                webRoot = Path.Combine(AppContext.BaseDirectory, "wwwroot");
            }

            var relativeDir = Path.Combine("uploads", "attachments", (entityType ?? "misc").ToLowerInvariant());
            var fullDir = Path.Combine(webRoot, relativeDir);

            if (!Directory.Exists(fullDir))
            {
                Directory.CreateDirectory(fullDir);
            }

            var physicalFilePath = Path.Combine(fullDir, storedFileName);
            var webRelativeUrl = $"/uploads/attachments/{(entityType ?? "misc").ToLowerInvariant()}/{storedFileName}";

            // Stream file to disk
            await using (var fileStream = new FileStream(physicalFilePath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
            await using (var readStream = file.OpenReadStream(maxAllowedSize: maxFileSize))
            {
                await readStream.CopyToAsync(fileStream, 81920);
            }

            var dto = new EntityAttachmentDto
            {
                EntityType = entityType,
                EntityId = entityId,
                EntitySecondaryId = entitySecondaryId,
                OriginalFileName = safeName,
                StoredFileName = storedFileName,
                FilePath = webRelativeUrl,
                FileSize = file.Size,
                FileExtension = extension,
                ContentType = string.IsNullOrWhiteSpace(file.ContentType) ? "application/octet-stream" : file.ContentType,
                Remarks = remarks,
                UploadedBy = string.IsNullOrWhiteSpace(uploadedBy) ? "System" : uploadedBy,
                UploadedAt = DateTime.Now,
                IsDeleted = false
            };

            var newId = await _attachmentDac.InsertAttachmentAsync(dto);
            dto.Id = newId;

            return dto;
        }

        public async Task<bool> DeleteAttachmentAsync(long id, string deletedBy = null)
        {
            try
            {
                return await _attachmentDac.SoftDeleteAttachmentAsync(id, deletedBy ?? "System");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error soft-deleting attachment ID {Id}", id);
                return false;
            }
        }

        public Task<bool> UpdateRemarksAsync(long id, string remarks)
        {
            // Future extension if inline remarks editing is desired
            return Task.FromResult(true);
        }

        public async Task<int> GetCountAsync(string entityType, string entityId, string entitySecondaryId = null)
        {
            try
            {
                return await _attachmentDac.GetAttachmentCountAsync(entityType, entityId, entitySecondaryId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting attachment count for {EntityType} {EntityId}", entityType, entityId);
                return 0;
            }
        }

        public string GetPhysicalFilePath(EntityAttachmentDto attachment)
        {
            if (attachment == null || string.IsNullOrWhiteSpace(attachment.StoredFileName))
                return string.Empty;

            var webRoot = _env.WebRootPath;
            if (string.IsNullOrWhiteSpace(webRoot))
            {
                webRoot = Path.Combine(AppContext.BaseDirectory, "wwwroot");
            }

            var relativeDir = Path.Combine("uploads", "attachments", (attachment.EntityType ?? "misc").ToLowerInvariant());
            return Path.Combine(webRoot, relativeDir, attachment.StoredFileName);
        }
    }
}
