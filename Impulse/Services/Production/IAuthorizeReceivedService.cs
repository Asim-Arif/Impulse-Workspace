using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IAuthorizeReceivedService
    {
        Task<List<PendingAuthorizeReceivedItemModel>> GetPendingAuthorizeReceivedItemsAsync(AuthorizeReceivedFilterModel filter);
        Task<List<WastageTypeLookupModel>> GetWastageTypesAsync();
        Task<List<RepairTypeLookupModel>> GetRepairTypesAsync();
        Task<List<EmployeeLookupModel>> GetInspectorsAsync();
        Task<List<ProcessPOLookupModel>> GetProcessesAsync();
        Task<List<MakerPOLookupModel>> GetMakersAsync();
        Task<bool> SaveAuthorizationAsync(List<PendingAuthorizeReceivedItemModel> items, string inspectorEmpId, string userName, string machineName, int userId);
    }
}
