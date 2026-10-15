using System;

namespace DataAccessLibrary.Models.ViewModels.Attachments
{
    public class EntityNoteDto
    {
        public long Id { get; set; }
        public string EntityType { get; set; } = string.Empty;
        public string EntityId { get; set; } = string.Empty;
        public string? EntitySecondaryId { get; set; }
        public string NoteText { get; set; } = string.Empty;
        public string LastModifiedBy { get; set; } = string.Empty;
        public DateTime LastModifiedAt { get; set; }
    }
}
