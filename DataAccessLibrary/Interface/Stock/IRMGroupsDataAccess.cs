using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IRMGroupsDataAccess
    {
        Task<List<RMGroupViewModel>> GetRMGroupsAsync();
        Task<int> InsertRMGroupAsync(RMGroupViewModel group);
        Task<int> UpdateRMGroupAsync(RMGroupViewModel group);
        Task<int> DeleteRMGroupAsync(int id);
        Task<int> CountLinkedMaterialsAsync(int groupId);
    }
}
