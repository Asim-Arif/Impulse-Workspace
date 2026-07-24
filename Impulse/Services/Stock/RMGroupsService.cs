using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class RMGroupsService : IRMGroupsService
    {
        private readonly IRMGroupsDataAccess _dataAccess;

        public RMGroupsService(IRMGroupsDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<RMGroupViewModel>> GetRMGroupsAsync()
        {
            return await _dataAccess.GetRMGroupsAsync();
        }

        public async Task<int> InsertRMGroupAsync(RMGroupViewModel group)
        {
            return await _dataAccess.InsertRMGroupAsync(group);
        }

        public async Task<int> UpdateRMGroupAsync(RMGroupViewModel group)
        {
            return await _dataAccess.UpdateRMGroupAsync(group);
        }

        public async Task<int> DeleteRMGroupAsync(int id)
        {
            return await _dataAccess.DeleteRMGroupAsync(id);
        }

        public async Task<int> CountLinkedMaterialsAsync(int groupId)
        {
            return await _dataAccess.CountLinkedMaterialsAsync(groupId);
        }
    }
}
