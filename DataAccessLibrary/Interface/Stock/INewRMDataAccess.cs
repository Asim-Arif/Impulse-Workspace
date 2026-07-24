using DataAccessLibrary.Models.ViewModels.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface INewRMDataAccess
    {
        Task<NewRMViewModel> GetRMAsync(int rmId);
        Task<int> InsertRMAsync(NewRMViewModel model);
        Task<bool> UpdateRMAsync(NewRMViewModel model);
        Task<bool> DeleteRMAsync(int id);
        Task<List<ItemLookupModel>> GetAvailableItemsAsync();
        Task<List<RawMaterialLookupModel>> GetAvailableRawMaterialsAsync();
    }
}
