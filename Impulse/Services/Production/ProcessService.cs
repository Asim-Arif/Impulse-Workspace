using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public class ProcessService : IProcessService
    {
        private readonly IProcessDataAccess _dataAccess;
        private readonly IAuditService _auditService;

        public ProcessService(IProcessDataAccess dataAccess, IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
        }

        public Task<List<ProcessListItem>> GetProcessesAsync(bool isInspectionProcess = false) => _dataAccess.GetProcessesAsync(isInspectionProcess);

        public Task<ProcessModel?> GetProcessByIdAsync(int processId) => _dataAccess.GetProcessByIdAsync(processId);

        public Task<ProcessLookups> GetLookupsAsync() => _dataAccess.GetLookupsAsync();

        public async Task<int> SaveProcessAsync(ProcessModel model)
        {
            if (string.IsNullOrWhiteSpace(model.Description))
                throw new InvalidOperationException("Please enter a valid Process Name.");

            return await _dataAccess.SaveProcessAsync(model);
        }

        public async Task<bool> UpdateProcessAsync(ProcessModel model)
        {
            if (string.IsNullOrWhiteSpace(model.Description))
                throw new InvalidOperationException("Please enter a valid Process Name.");

            return await _dataAccess.UpdateProcessAsync(model);
        }

        public Task<(bool CanDelete, string Reason)> CanDeleteProcessAsync(int processId) => _dataAccess.CanDeleteProcessAsync(processId);

        public Task<bool> DeleteProcessAsync(int processId) => _dataAccess.DeleteProcessAsync(processId);

        public Task<List<ProcessInspectionPointModel>> GetInspectionPointsAsync(int processId) => _dataAccess.GetInspectionPointsAsync(processId);

        public async Task<bool> AddInspectionPointAsync(ProcessInspectionPointModel point)
        {
            if (string.IsNullOrWhiteSpace(point.Point_Description))
                throw new InvalidOperationException("Please enter an inspection checkpoint description.");

            return await _dataAccess.AddInspectionPointAsync(point);
        }

        public async Task<bool> UpdateInspectionPointAsync(ProcessInspectionPointModel point)
        {
            if (string.IsNullOrWhiteSpace(point.Point_Description))
                throw new InvalidOperationException("Please enter an inspection checkpoint description.");

            return await _dataAccess.UpdateInspectionPointAsync(point);
        }

        public Task<bool> DeleteInspectionPointAsync(int entryId) => _dataAccess.DeleteInspectionPointAsync(entryId);
    }
}
