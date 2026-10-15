using System;
using System.IO;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using Impulse.Services.Common;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;

namespace Impulse.Controllers
{
    [Route("api/attachments")]
    [ApiController]
    public class AttachmentsController : ControllerBase
    {
        private readonly IEntityAttachmentDataAccess _attachmentDac;
        private readonly IEntityAttachmentService _attachmentService;
        private readonly IWebHostEnvironment _env;

        public AttachmentsController(
            IEntityAttachmentDataAccess attachmentDac,
            IEntityAttachmentService attachmentService,
            IWebHostEnvironment env)
        {
            _attachmentDac = attachmentDac;
            _attachmentService = attachmentService;
            _env = env;
        }

        [HttpGet("download/{id:long}")]
        public async Task<IActionResult> Download(long id)
        {
            var attachment = await _attachmentDac.GetAttachmentByIdAsync(id);
            if (attachment == null || attachment.IsDeleted)
            {
                return NotFound("Attachment not found or has been deleted.");
            }

            var physicalPath = _attachmentService.GetPhysicalFilePath(attachment);
            if (!System.IO.File.Exists(physicalPath))
            {
                return NotFound("Physical file not found on disk.");
            }

            var contentType = attachment.ContentType;
            if (string.IsNullOrWhiteSpace(contentType) || contentType == "application/octet-stream")
            {
                var provider = new FileExtensionContentTypeProvider();
                if (!provider.TryGetContentType(attachment.OriginalFileName, out contentType))
                {
                    contentType = "application/octet-stream";
                }
            }

            return PhysicalFile(physicalPath, contentType, attachment.OriginalFileName, enableRangeProcessing: true);
        }

        [HttpGet("view/{id:long}")]
        public async Task<IActionResult> ViewFile(long id)
        {
            var attachment = await _attachmentDac.GetAttachmentByIdAsync(id);
            if (attachment == null || attachment.IsDeleted)
            {
                return NotFound("Attachment not found or has been deleted.");
            }

            var physicalPath = _attachmentService.GetPhysicalFilePath(attachment);
            if (!System.IO.File.Exists(physicalPath))
            {
                return NotFound("Physical file not found on disk.");
            }

            var contentType = attachment.ContentType;
            var provider = new FileExtensionContentTypeProvider();
            if (string.IsNullOrWhiteSpace(contentType) || contentType == "application/octet-stream")
            {
                if (!provider.TryGetContentType(attachment.OriginalFileName, out contentType))
                {
                    contentType = "application/octet-stream";
                }
            }

            Response.Headers["Content-Disposition"] = $"inline; filename=\"{Uri.EscapeDataString(attachment.OriginalFileName)}\"";
            return PhysicalFile(physicalPath, contentType, enableRangeProcessing: true);
        }
    }
}
