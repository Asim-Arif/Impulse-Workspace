using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.QMS
{
    public class QMSReasonModel
    {
        public int EntryID { get; set; }

        [Required(ErrorMessage = "Reason Description is required.")]
        [StringLength(50, ErrorMessage = "Reason Description cannot exceed 50 characters.")]
        public string Reason_QMS { get; set; } = string.Empty;

        public int UsageCount { get; set; }

        // UI state helpers for inline editing
        public bool IsEditing { get; set; }
        public string EditBuffer { get; set; } = string.Empty;
    }
}
