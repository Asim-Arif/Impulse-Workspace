using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public class ProcessGroupService : IProcessGroupService
    {
        private readonly IProcessGroupDataAccess _dataAccess;
        private readonly IAuditService _auditService;

        public ProcessGroupService(IProcessGroupDataAccess dataAccess, IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
        }

        public Task<List<ProcessGroupListItem>> GetProcessGroupsAsync() => _dataAccess.GetProcessGroupsAsync();

        public Task<ProcessGroupDto?> GetProcessGroupByIdAsync(int entryId) => _dataAccess.GetProcessGroupByIdAsync(entryId);

        public async Task<int> SaveProcessGroupAsync(ProcessGroupDto model)
        {
            if (string.IsNullOrWhiteSpace(model.GroupName))
                throw new InvalidOperationException("Please enter a Group Name/Description.");

            if (model.Processes == null || !model.Processes.Any())
                throw new InvalidOperationException("Please add at least one Process to this Group sequence.");

            return await _dataAccess.SaveProcessGroupAsync(model);
        }

        public async Task<bool> UpdateProcessGroupAsync(ProcessGroupDto model)
        {
            if (string.IsNullOrWhiteSpace(model.GroupName))
                throw new InvalidOperationException("Please enter a Group Name/Description.");

            if (model.Processes == null || !model.Processes.Any())
                throw new InvalidOperationException("Please add at least one Process to this Group sequence.");

            return await _dataAccess.UpdateProcessGroupAsync(model);
        }

        public Task<(bool CanDelete, string Reason)> CanDeleteProcessGroupAsync(int entryId) => _dataAccess.CanDeleteProcessGroupAsync(entryId);

        public Task<bool> DeleteProcessGroupAsync(int entryId) => _dataAccess.DeleteProcessGroupAsync(entryId);

        public Task<bool> UpdateItemProcessesFromProcessGroupAsync(int entryId) => _dataAccess.UpdateItemProcessesFromProcessGroupAsync(entryId);

        public Task<List<LookupItemInt>> GetAvailableProcessesLookupAsync() => _dataAccess.GetAvailableProcessesLookupAsync();
    }
}
