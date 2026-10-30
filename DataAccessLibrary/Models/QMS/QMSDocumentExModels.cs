using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.QMS
{
    public class QMSDocumentExListItemDto
    {
        public int EntryID { get; set; }
        public string? DocumentNumber { get; set; }
        public string? DocumentName { get; set; }
        public string? Revision { get; set; }
        public string? DocumentType { get; set; }
        public string? Master_DocumentType { get; set; }
        public DateTime? IssueDT { get; set; }
        public string? InitiaterEmpID { get; set; }
        public string? InitiatorName { get; set; }
        public string? DeptName { get; set; }
        public DateTime? LastApprovedEntryDT { get; set; }
        public bool IsSelected { get; set; }
    }

    public class QMSDocumentExModel
    {
        public int EntryID { get; set; }

        [Required(ErrorMessage = "Document Number is required.")]
        [StringLength(50, ErrorMessage = "Document Number cannot exceed 50 characters.")]
        public string DocumentNumber { get; set; } = string.Empty;

        [Required(ErrorMessage = "Document Title/Name is required.")]
        [StringLength(100, ErrorMessage = "Document Title cannot exceed 100 characters.")]
        public string DocumentName { get; set; } = string.Empty;

        [StringLength(50, ErrorMessage = "Revision cannot exceed 50 characters.")]
        public string? Revision { get; set; } = "00";

        [StringLength(50)]
        public string? DocumentType { get; set; }

        [Required(ErrorMessage = "Category is required.")]
        [StringLength(50)]
        public string Master_DocumentType { get; set; } = string.Empty;

        public DateTime? IssueDT { get; set; } = DateTime.Today;

        public string? InitiaterEmpID { get; set; }
        public string? InitiatorName { get; set; }
        public string? InitiatorDept { get; set; }

        public bool Internal { get; set; } = true;
    }

    public class QMSDocumentExDistributionItemModel
    {
        public int EntryID { get; set; }
        public int QMS_RefID { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public string? Name { get; set; }
        public string? DeptName { get; set; }
        public string? Designation { get; set; }
        public bool IsNew { get; set; }
        public bool IsDeleted { get; set; }
    }

    public class QMSMasterDocTypeOption
    {
        public int EntryID { get; set; }
        public string MasterDocumentType { get; set; } = string.Empty;
        public string DocumentType { get; set; } = string.Empty;
    }
}

