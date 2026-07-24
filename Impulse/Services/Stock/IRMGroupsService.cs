using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IRMGroupsService
    {
        Task<List<RMGroupViewModel>> GetRMGroupsAsync();
        Task<int> InsertRMGroupAsync(RMGroupViewModel group);
        Task<int> UpdateRMGroupAsync(RMGroupViewModel group);
        Task<int> DeleteRMGroupAsync(int id);
        Task<int> CountLinkedMaterialsAsync(int groupId);
    }
}
