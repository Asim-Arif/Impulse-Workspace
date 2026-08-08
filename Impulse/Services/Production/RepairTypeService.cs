using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public class RepairTypeService : IRepairTypeService
    {
        private readonly IRepairTypeDataAccess _dataAccess;
        private readonly IAuditService _auditService;

        public RepairTypeService(IRepairTypeDataAccess dataAccess, IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
        }

        public Task<List<RepairTypeListItem>> GetRepairTypesAsync() => _dataAccess.GetRepairTypesAsync();

        public Task<RepairTypeDto?> GetRepairTypeByIdAsync(int entryId) => _dataAccess.GetRepairTypeByIdAsync(entryId);

        public Task<List<RepairTypeProcessItemDto>> GetAvailableProcessesForRepairAsync(int? repairEntryId = null) =>
            _dataAccess.GetAvailableProcessesForRepairAsync(repairEntryId);

        public async Task<int> SaveRepairTypeAsync(RepairTypeDto model)
        {
            Validate(model);
            return await _dataAccess.SaveRepairTypeAsync(model);
        }

        public async Task<bool> UpdateRepairTypeAsync(RepairTypeDto model)
        {
            Validate(model);
            return await _dataAccess.UpdateRepairTypeAsync(model);
        }

        public Task<(bool CanDelete, string Reason)> CanDeleteRepairTypeAsync(int entryId) => _dataAccess.CanDeleteRepairTypeAsync(entryId);

        public Task<bool> DeleteRepairTypeAsync(int entryId) => _dataAccess.DeleteRepairTypeAsync(entryId);

        private static void Validate(RepairTypeDto model)
        {
            if (string.IsNullOrWhiteSpace(model.RepairCode))
                throw new InvalidOperationException("Please enter a Repair Code (e.g. REP-01, DEF-CUT).");

            if (string.IsNullOrWhiteSpace(model.RepairType))
                throw new InvalidOperationException("Please enter a Repair Type description.");

            if (string.IsNullOrWhiteSpace(model.DefectType))
                throw new InvalidOperationException("Please enter a Defect Type.");

            if (model.Processes == null || !model.Processes.Any(p => p.IsSelected))
                throw new InvalidOperationException("Please select at least one manufacturing process for this Repair Type.");
        }
    }
}
