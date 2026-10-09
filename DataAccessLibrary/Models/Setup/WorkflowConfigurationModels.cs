using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.Setup
{
    public class WorkflowConfigurationModel
    {
        public string WorkflowCode { get; set; } = string.Empty;
        public string WorkflowName { get; set; } = string.Empty;
        public string Module { get; set; } = string.Empty;
        public string? Description { get; set; }
        public bool IsEnabled { get; set; } = true;
        public string ApproverRole { get; set; } = "Director";
        public DateTime UpdatedDate { get; set; } = DateTime.UtcNow;
        public string? UpdatedBy { get; set; }

        public List<string> ExemptRoles { get; set; } = new List<string>();
    }

    public class WorkflowExemptRoleModel
    {
        public int ID { get; set; }
        public string WorkflowCode { get; set; } = string.Empty;
        public string RoleName { get; set; } = string.Empty;
    }
}
