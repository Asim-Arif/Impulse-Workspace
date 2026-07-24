using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface INewRMService
    {
        Task<NewRMViewModel> GetRMAsync(int rmId);
        Task<int> SaveRMAsync(NewRMViewModel model);
        Task<bool> UpdateRMAsync(NewRMViewModel model);
        Task<List<ItemLookupModel>> GetAvailableItemsAsync();
        Task<List<RawMaterialLookupModel>> GetAvailableRawMaterialsAsync();
    }
}
