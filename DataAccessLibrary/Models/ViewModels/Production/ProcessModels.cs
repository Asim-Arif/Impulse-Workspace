using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class ProcessModel
    {
        public int ProcessID { get; set; }
        public int SNo { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty; // Process Name
        public string Supervisor { get; set; } = string.Empty; // Detailed Description
        public int Operation { get; set; } = 0; // 0: Employee, 1: Contractor, 2: Both
        public int AuthRequired { get; set; } = 0; // 0: No, 1: Yes
        public string ProcessNameUrdu { get; set; } = string.Empty;
        public string ProcessNameUrduOther { get; set; } = string.Empty;
        public int? Insp_RefID { get; set; }
        public int? Fix_Maker_RefID { get; set; }
        public bool InspectionProcess { get; set; } = false;
        public int? BillingProcessID { get; set; }
        public bool IsSelected { get; set; }
    }

    public class ProcessListItem
    {
        public int ProcessID { get; set; }
        public int SNO { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Supervisor { get; set; } = string.Empty;
        public int Operation { get; set; }
        public int? AuthRequired { get; set; }
        public string ProcessNameUrdu { get; set; } = string.Empty;
        public string ProcessNameUrduOther { get; set; } = string.Empty;
        public int? Insp_RefID { get; set; }
        public int? Fix_Maker_RefID { get; set; }
        public string VenderName { get; set; } = string.Empty;
        public string Insp_Code { get; set; } = string.Empty;
        public string Insp_ProcessName { get; set; } = string.Empty;
        public string Insp_ProcessNameUrdu { get; set; } = string.Empty;
        public bool? InspectionProcess { get; set; }
        public int? BillingProcessID { get; set; }
        public string BillingProcessName { get; set; } = string.Empty;

        public string OperationName => Operation switch
        {
            0 => "Employee",
            1 => "Contractor",
            2 => "Both",
            _ => "Employee"
        };

        public bool IsAuthRequired => (AuthRequired ?? 0) == 1;
    }

    public class ProcessInspectionPointModel
    {
        public int EntryID { get; set; }
        public int ProcessID { get; set; }
        public string Point_Description { get; set; } = string.Empty;
    }

    public class ProcessLookups
    {
        public List<LookupItemInt> Makers { get; set; } = new();
        public List<LookupItemInt> InspectionProcesses { get; set; } = new();
        public List<LookupItemInt> BillingProcesses { get; set; } = new();
    }
}
