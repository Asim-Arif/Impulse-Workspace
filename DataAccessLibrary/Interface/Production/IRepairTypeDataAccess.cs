using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IRepairTypeDataAccess
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
