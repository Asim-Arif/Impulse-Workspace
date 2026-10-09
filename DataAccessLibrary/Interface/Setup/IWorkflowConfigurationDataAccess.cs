using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace DataAccessLibrary.Interface.Setup
{
    public interface IWorkflowConfigurationDataAccess
    {
        Task<List<WorkflowConfigurationModel>> GetAllWorkflowConfigsAsync();
        Task<WorkflowConfigurationModel?> GetWorkflowConfigByCodeAsync(string workflowCode);
        Task<bool> UpdateWorkflowConfigAsync(WorkflowConfigurationModel model, string updatedBy);
        Task<bool> IsWorkflowApprovalRequiredAsync(string workflowCode, IEnumerable<string> userRoles);
    }
}
