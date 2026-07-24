using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class NewRMService : INewRMService
    {
        private readonly INewRMDataAccess _newRMDataAccess;

        public NewRMService(INewRMDataAccess newRMDataAccess)
        {
            _newRMDataAccess = newRMDataAccess;
        }

        public async Task<bool> DeleteRMAsync(int id)
        {
            return await _newRMDataAccess.DeleteRMAsync(id);
        }

        public async Task<List<ItemLookupModel>> GetAvailableItemsAsync()
        {
            return await _newRMDataAccess.GetAvailableItemsAsync();
        }

        public async Task<List<RawMaterialLookupModel>> GetAvailableRawMaterialsAsync()
        {
            return await _newRMDataAccess.GetAvailableRawMaterialsAsync();
        }

        public Task<NewRMViewModel> GetRMAsync(int rmId)
        {
            return _newRMDataAccess.GetRMAsync(rmId);
        }

        public async Task<int> SaveRMAsync(NewRMViewModel model)
        {
            if (model.RMID == 0)
            {
                return await _newRMDataAccess.InsertRMAsync(model);
            }
            else
            {
                await _newRMDataAccess.UpdateRMAsync(model);
                return model.RMID;
            }
        }

        public async Task<bool> UpdateRMAsync(NewRMViewModel model)
        {
            return await _newRMDataAccess.UpdateRMAsync(model);
        }
    }
}
