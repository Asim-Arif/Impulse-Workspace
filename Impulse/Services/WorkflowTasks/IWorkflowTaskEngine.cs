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
        Task<bool> RejectRoleTaskAsync(string entityType, string entityRefId, string targetRole, string rejectedByUserName, string? rejectionReason = null);
        Task<bool> AuthorizeCustomerOrderAsync(string orderNo, string authorizedByUserName);

        // Voucher Deletion Task Workflow
        Task<int> RequestVoucherDeletionAsync(string vchrNo, string originatorUserName, string deleteReason, string? machineName = null);
        Task<bool> ApproveVoucherDeletionAsync(string vchrNo, string approvedByDirectorUserName, string? directorRemarks = null);
        Task<bool> RejectVoucherDeletionAsync(string vchrNo, string rejectedByDirectorUserName, string rejectionReason);
        Task<List<DataAccessLibrary.Models.IntraOffice.TaskItem>> GetPendingVoucherDeletionTasksAsync();

        // Attendance Approval Workflow
        Task<int> RequestAttendanceActionAsync(DataAccessLibrary.Models.ViewModels.Payroll.AttendanceWorkflowRequestDto request);
        Task<bool> ApproveAttendanceActionAsync(int taskId, string approvedByDirectorUserName, string? directorRemarks = null);
        Task<bool> RejectAttendanceActionAsync(int taskId, string rejectedByDirectorUserName, string rejectionReason);
        Task<List<DataAccessLibrary.Models.IntraOffice.TaskItem>> GetPendingAttendanceTasksAsync();
        Task<DataAccessLibrary.Models.ViewModels.Payroll.AttendanceWorkflowRequestDto?> GetAttendanceTaskDetailsAsync(int taskId);

        // Maker Item Rate Approval Workflow
        Task<int> RequestMakerItemRateChangeAsync(DataAccessLibrary.Models.ViewModels.Production.MakerItemRateWorkflowRequestDto request);
        Task<bool> ApproveMakerItemRateActionAsync(int taskId, string approvedByDirectorUserName, string? directorRemarks = null);
        Task<bool> RejectMakerItemRateActionAsync(int taskId, string rejectedByDirectorUserName, string rejectionReason);
        Task<List<DataAccessLibrary.Models.IntraOffice.TaskItem>> GetPendingMakerRateTasksAsync();
        Task<DataAccessLibrary.Models.ViewModels.Production.MakerItemRateWorkflowRequestDto?> GetMakerRateTaskDetailsAsync(int taskId);

        // Production Deletion Requests Workflow (Lot Receiving, Issuance, etc.)
        Task<int> RequestLotReceivingDeletionAsync(DataAccessLibrary.Models.ViewModels.Production.ProductionDeletionRequestModel request);
        Task<bool> ApproveLotReceivingDeletionAsync(int requestId, string approvedByDirectorUserName, string? directorRemarks = null);
        Task<bool> RejectLotReceivingDeletionAsync(int requestId, string rejectedByDirectorUserName, string rejectionReason);
        Task<List<DataAccessLibrary.Models.ViewModels.Production.ProductionDeletionRequestModel>> GetPendingLotDeletionRequestsAsync();
        Task<DataAccessLibrary.Models.ViewModels.Production.ProductionDeletionRequestModel?> GetLotDeletionRequestDetailsAsync(int requestId);

        // Production Issuance Deletion Workflow
        Task<int> RequestProductionIssuanceDeletionAsync(DataAccessLibrary.Models.ViewModels.Production.ProductionDeletionRequestModel request);
        Task<bool> ApproveProductionIssuanceDeletionAsync(int requestId, string approvedByDirectorUserName, string? directorRemarks = null);
        Task<bool> RejectProductionIssuanceDeletionAsync(int requestId, string rejectedByDirectorUserName, string rejectionReason);
        Task<List<DataAccessLibrary.Models.ViewModels.Production.ProductionDeletionRequestModel>> GetPendingIssuanceDeletionRequestsAsync();

        // Skip Process Governance Workflow
        Task<int> RequestSkipProcessAsync(DataAccessLibrary.Models.ViewModels.Production.ProductionDeletionRequestModel request);
        Task<bool> ApproveSkipProcessAsync(int requestId, string approvedByDirectorUserName, string? directorRemarks = null);
        Task<bool> RejectSkipProcessAsync(int requestId, string rejectedByDirectorUserName, string rejectionReason);
        Task<List<DataAccessLibrary.Models.ViewModels.Production.ProductionDeletionRequestModel>> GetPendingSkipProcessRequestsAsync();

        // Governance Workflow Config Checks
        Task<bool> IsApprovalRequiredAsync(string workflowCode, string userName, System.Security.Claims.ClaimsPrincipal? claimsPrincipal = null);
        Task<DataAccessLibrary.Models.Setup.WorkflowConfigurationModel?> GetWorkflowConfigAsync(string workflowCode);
    }
}
