using System;

namespace DataAccessLibrary.Models.ViewModels.Attachments
{
    public class EntityAttachmentDto
    {
        public long Id { get; set; }
        public string EntityType { get; set; } = string.Empty;
        public string EntityId { get; set; } = string.Empty;
        public string? EntitySecondaryId { get; set; }
        public string OriginalFileName { get; set; } = string.Empty;
        public string StoredFileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string FileExtension { get; set; } = string.Empty;
        public string? ContentType { get; set; }
        public string? Remarks { get; set; }
        public string UploadedBy { get; set; } = string.Empty;
        public DateTime UploadedAt { get; set; }
        public bool IsDeleted { get; set; }

        public string FormattedFileSize
        {
            get
            {
                if (FileSize < 1024) return $"{FileSize} B";
                if (FileSize < 1024 * 1024) return $"{FileSize / 1024.0:F1} KB";
                return $"{FileSize / (1024.0 * 1024.0):F2} MB";
            }
        }

        public string FileIconClass
        {
            get
            {
                var ext = (FileExtension ?? string.Empty).ToLowerInvariant().TrimStart('.');
                return ext switch
                {
                    "pdf" => "fas fa-file-pdf text-danger",
                    "doc" or "docx" => "fas fa-file-word text-primary",
                    "xls" or "xlsx" or "csv" => "fas fa-file-excel text-success",
                    "ppt" or "pptx" => "fas fa-file-powerpoint text-warning",
                    "jpg" or "jpeg" or "png" or "gif" or "bmp" or "webp" or "svg" => "fas fa-file-image text-info",
                    "zip" or "rar" or "7z" or "tar" or "gz" => "fas fa-file-zipper text-warning",
                    "txt" or "log" or "md" => "fas fa-file-lines text-secondary",
                    "json" or "xml" or "html" => "fas fa-file-code text-secondary",
                    "mp3" or "wav" or "ogg" => "fas fa-file-audio text-info",
                    "mp4" or "avi" or "mkv" or "mov" => "fas fa-file-video text-danger",
                    _ => "fas fa-file text-muted"
                };
            }
        }

        public bool CanPreviewInline
        {
            get
            {
                var ext = (FileExtension ?? string.Empty).ToLowerInvariant().TrimStart('.');
                return ext is "pdf" or "jpg" or "jpeg" or "png" or "gif" or "webp" or "svg" or "txt";
            }
        }

        public string FileBadgeColor => string.Empty;
    }
}
