using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.WorkflowTasks
{
    public class WorkflowTaskCreateRequest
    {
        public string SourceEntityType { get; set; } = string.Empty;
        public string SourceEntityRefId { get; set; } = string.Empty;
        
        public string? TargetRole { get; set; }
        public List<string>? TargetRoles { get; set; }

        public List<string>? TargetUserNames { get; set; }

        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string ActionUrl { get; set; } = string.Empty;
        public int Priority { get; set; } = 2;
        public DateTime? DueDate { get; set; }
        public string CreatedBy { get; set; } = "System";
    }

    public interface IWorkflowTaskEngine
    {
        Task<int> CreateRoleTaskAsync(WorkflowTaskCreateRequest request);
        Task<bool> CompleteTaskAsync(string entityType, string entityRefId, string completedByUserName, string? notes = null);
        Task<bool> CompleteRoleTaskAsync(string entityType, string entityRefId, string targetRole, string completedByUserName, string? notes = null);
        Task<bool> AuthorizeCustomerOrderAsync(string orderNo, string authorizedByUserName);
    }
}
