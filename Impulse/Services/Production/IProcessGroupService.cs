using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IProcessGroupService
    {
        Task<List<ProcessGroupListItem>> GetProcessGroupsAsync();
        Task<ProcessGroupDto?> GetProcessGroupByIdAsync(int entryId);
        Task<int> SaveProcessGroupAsync(ProcessGroupDto model);
        Task<bool> UpdateProcessGroupAsync(ProcessGroupDto model);
        Task<(bool CanDelete, string Reason)> CanDeleteProcessGroupAsync(int entryId);
        Task<bool> DeleteProcessGroupAsync(int entryId);
        Task<bool> UpdateItemProcessesFromProcessGroupAsync(int entryId);
        Task<List<LookupItemInt>> GetAvailableProcessesLookupAsync();
    }
}
