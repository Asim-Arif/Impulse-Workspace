using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.QMS
{
    public class QMSDocumentListItemModel
    {
        public int EntryID { get; set; }
        public string? DocumentNumber { get; set; }
        public string? DocumentName { get; set; }
        public string? Revision { get; set; }
        public string? DocumentType { get; set; }
        public DateTime? IssueDT { get; set; }
        public bool? Internal { get; set; }
        public string? InitiaterEmpID { get; set; }
        public string? InitiatorName { get; set; }
        public string? DeptName { get; set; }
        public int DistributionCount { get; set; }
        public int AttachmentCount { get; set; }
        public bool IsSelected { get; set; }
    }

    public class QMSDocumentModel
    {
        public int EntryID { get; set; }
        public string? DocumentNumber { get; set; }
        public string? DocumentName { get; set; }
        public string? Revision { get; set; } = "00";
        public string? DocumentType { get; set; }
        public DateTime? IssueDT { get; set; } = DateTime.Today;
        public string? InitiaterEmpID { get; set; }
        public string? InitiatorName { get; set; }
        public string? InitiatorDepartment { get; set; }
        public string? InitiatorDesignation { get; set; }
        public bool Internal { get; set; } = true;
        public List<QMSDocumentDistributionModel> DistributionList { get; set; } = new();
    }

    public class QMSDocumentDistributionModel
    {
        public int EntryID { get; set; }
        public int QMS_RefID { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public string? EmpName { get; set; }
        public string? DeptName { get; set; }
        public string? Designation { get; set; }
        public bool IsMarkedForDeletion { get; set; }
    }

    public class QMSDocumentFilterDto
    {
        public string? SearchText { get; set; }
        public string? DocumentType { get; set; }
        public int? InternalFilter { get; set; } // null/all, 1 = Internal, 0 = External
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
    }

    public class EmployeeDropdownDto
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? FName { get; set; }
        public string? DeptName { get; set; }
        public string? Designation { get; set; }

        public string DisplayText => $"[{EmpID}] {Name}" + (!string.IsNullOrWhiteSpace(Designation) ? $" ({Designation})" : "");
    }
}

