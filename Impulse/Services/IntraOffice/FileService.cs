using System;
using System.IO;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;

namespace Impulse.Services.IntraOffice
{
    public interface IFileService
    {
        Task<MessageAttachment> UploadFileAsync(IBrowserFile file, long messageId);
        Task<string> SaveFileToDiskAsync(IBrowserFile file, long? maxAllowedSize = null);
        Task<string> SaveBase64FileToDiskAsync(string base64Data, string fileName);
        Task<MessageAttachment?> GetAttachmentAsync(int id);
        Task DeleteFileAsync(int attachmentId);
        string GetUploadPath();
        string GetAbsolutePath(string relativePath);
    }

    public class FileService : IFileService
    {
        private readonly IWebHostEnvironment _environment;
        private readonly IConfiguration _configuration;
        private const long DefaultMaxFileSize = 500L * 1024 * 1024; // 500 MB default

        public FileService(IWebHostEnvironment environment, IConfiguration configuration)
        {
            _environment = environment;
            _configuration = configuration;
        }

        private long GetConfiguredMaxFileSize()
        {
            var configuredMb = _configuration.GetValue<long?>("FileUpload:MaxFileSizeInMB");
            if (configuredMb.HasValue && configuredMb.Value > 0)
            {
                return configuredMb.Value * 1024L * 1024L;
            }
            return DefaultMaxFileSize;
        }

        public async Task<string> SaveFileToDiskAsync(IBrowserFile file, long? maxAllowedSize = null)
        {
            var uploadsPath = GetUploadPath();
            Directory.CreateDirectory(uploadsPath);
            var safeName = Path.GetFileName(file.Name);
            var uniqueFileName = $"{Guid.NewGuid()}_{safeName}";
            var filePath = Path.Combine(uploadsPath, uniqueFileName);

            var effectiveLimit = maxAllowedSize ?? GetConfiguredMaxFileSize();
            await using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true);
            await using var readStream = file.OpenReadStream(effectiveLimit);
            await readStream.CopyToAsync(stream, 81920);
            return $"/uploads/{uniqueFileName}";
        }

        public async Task<string> SaveBase64FileToDiskAsync(string base64Data, string fileName)
        {
            var uploadsPath = GetUploadPath();
            Directory.CreateDirectory(uploadsPath);
            var safeName = Path.GetFileName(fileName);
            var uniqueFileName = $"{Guid.NewGuid()}_{safeName}";
            var filePath = Path.Combine(uploadsPath, uniqueFileName);
            var bytes = Convert.FromBase64String(base64Data);
            await File.WriteAllBytesAsync(filePath, bytes);
            return $"/uploads/{uniqueFileName}";
        }

        public async Task<MessageAttachment> UploadFileAsync(IBrowserFile file, long messageId)
        {
            var relativePath = await SaveFileToDiskAsync(file);
            return new MessageAttachment
            {
                MessageId = messageId,
                FileName = file.Name,
                FilePath = relativePath,
                FileSize = file.Size,
                ContentType = file.ContentType,
                UploadedAt = DateTime.UtcNow
            };
        }

        public Task<MessageAttachment?> GetAttachmentAsync(int id)
        {
            return Task.FromResult<MessageAttachment?>(null);
        }

        public Task DeleteFileAsync(int attachmentId)
        {
            return Task.CompletedTask;
        }

        public string GetUploadPath() => Path.Combine(_environment.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot"), "uploads");
        public string GetAbsolutePath(string relativePath) => Path.Combine(_environment.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot"), relativePath.TrimStart('/'));
    }
}
