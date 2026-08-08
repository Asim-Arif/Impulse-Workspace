using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class RepairTypeListItem
    {
        public int EntryID { get; set; }
        public string RepairCode { get; set; } = string.Empty;
        public string RepairType { get; set; } = string.Empty;
        public string DefectType { get; set; } = string.Empty;
        public bool ReturnToOriginator { get; set; }
        public int ProcessCount { get; set; }
        public string ProcessSummary { get; set; } = string.Empty;
        public List<RepairTypeProcessDto> Processes { get; set; } = new();
    }

    public class RepairTypeDto
    {
        public int EntryID { get; set; }
        public string RepairCode { get; set; } = string.Empty;
        public string RepairType { get; set; } = string.Empty;
        public string DefectType { get; set; } = string.Empty;
        public bool ReturnToOriginator { get; set; }
        public List<RepairTypeProcessItemDto> Processes { get; set; } = new();
    }

    public class RepairTypeProcessItemDto
    {
        public int ProcessID { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public bool IsSelected { get; set; }
        public int SeqNo { get; set; }
        public bool Scanning { get; set; }
    }

    public class RepairTypeProcessDto
    {
        public int EntryID { get; set; }
        public int Repair_RefID { get; set; }
        public int ProcessID { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int SeqNo { get; set; }
        public bool Scanning { get; set; }
    }
}
