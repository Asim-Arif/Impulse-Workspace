using System;

namespace DataAccessLibrary.Models.QMS
{
    public class QMSDocumentTypeModel
    {
        public int EntryID { get; set; }
        public string DocumentType { get; set; } = string.Empty;
        public int UsageCount { get; set; }

        // UI state helpers for inline editing
        public bool IsEditing { get; set; }
        public string EditBuffer { get; set; } = string.Empty;
    }
}

