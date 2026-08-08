using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IProcessService
    {
        Task<List<ProcessListItem>> GetProcessesAsync(bool isInspectionProcess = false);
        Task<ProcessModel?> GetProcessByIdAsync(int processId);
        Task<ProcessLookups> GetLookupsAsync();
        Task<int> SaveProcessAsync(ProcessModel model);
        Task<bool> UpdateProcessAsync(ProcessModel model);
        Task<(bool CanDelete, string Reason)> CanDeleteProcessAsync(int processId);
        Task<bool> DeleteProcessAsync(int processId);

        // Inspection Points
        Task<List<ProcessInspectionPointModel>> GetInspectionPointsAsync(int processId);
        Task<bool> AddInspectionPointAsync(ProcessInspectionPointModel point);
        Task<bool> UpdateInspectionPointAsync(ProcessInspectionPointModel point);
        Task<bool> DeleteInspectionPointAsync(int entryId);
    }
}
