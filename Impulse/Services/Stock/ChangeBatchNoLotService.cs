using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class ChangeBatchNoLotService : IChangeBatchNoLotService
    {
        private readonly IChangeBatchNoLotDataAccess _dataAccess;

        public ChangeBatchNoLotService(IChangeBatchNoLotDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ItemViewModel>> GetItemsAsync()
        {
            return _dataAccess.GetItemsAsync();
        }

        public Task<List<LotListViewModel>> GetLotsByMaterialAsync(string materialId)
        {
            return _dataAccess.GetLotsByMaterialAsync(materialId);
        }

        public Task UpdateLotsAsync(List<LotListViewModel> lots, string userName, string machineName)
        {
            return _dataAccess.UpdateLotsAsync(lots, userName, machineName);
        }
    }
}
