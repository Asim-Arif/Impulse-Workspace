using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class ProcessGroupListItem
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; } = string.Empty;
        public int ProcessCount { get; set; }
        public string FlowSummary { get; set; } = string.Empty;
        public List<ProcessGroupChildDto> Children { get; set; } = new();
    }

    public class ProcessGroupDto
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; } = string.Empty;
        public List<ProcessGroupChildDto> Processes { get; set; } = new();
    }

    public class ProcessGroupChildDto
    {
        public int EntryID { get; set; }
        public int Group_RefID { get; set; }
        public int Process_RefID { get; set; }
        public string ProcessCode { get; set; } = string.Empty;
        public string ProcessName { get; set; } = string.Empty;
        public bool Scanning { get; set; } = false;
        public string Hub_Name { get; set; } = string.Empty;
        public int SeqNo { get; set; }
    }
}
