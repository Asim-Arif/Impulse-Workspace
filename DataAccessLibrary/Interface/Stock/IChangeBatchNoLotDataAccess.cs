using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IChangeBatchNoLotDataAccess
    {
        Task<List<ItemViewModel>> GetItemsAsync();
        Task<List<LotListViewModel>> GetLotsByMaterialAsync(string materialId);
        Task UpdateLotsAsync(List<LotListViewModel> lots, string userName, string machineName);
    }
}
