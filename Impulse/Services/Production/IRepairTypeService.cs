using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IRepairTypeService
    {
        Task<List<RepairTypeListItem>> GetRepairTypesAsync();
        Task<RepairTypeDto?> GetRepairTypeByIdAsync(int entryId);
        Task<List<RepairTypeProcessItemDto>> GetAvailableProcessesForRepairAsync(int? repairEntryId = null);
        Task<int> SaveRepairTypeAsync(RepairTypeDto model);
        Task<bool> UpdateRepairTypeAsync(RepairTypeDto model);
        Task<(bool CanDelete, string Reason)> CanDeleteRepairTypeAsync(int entryId);
        Task<bool> DeleteRepairTypeAsync(int entryId);
    }
}
