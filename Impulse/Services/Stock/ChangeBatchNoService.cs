using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;
using DataAccessLibrary.Interface.Stock;

namespace Impulse.Services.Stock
{
    public class ChangeBatchNoService : IChangeBatchNoService
    {
        private readonly IChangeBatchNoDataAccess _changeBatchNoDataAccess;

        public ChangeBatchNoService(IChangeBatchNoDataAccess changeBatchNoDataAccess)
        {
            _changeBatchNoDataAccess = changeBatchNoDataAccess;
        }

        public async Task<IEnumerable<object>> GetMaterialsAsync()
        {
            return await _changeBatchNoDataAccess.GetMaterialsAsync();
        }

        public async Task<IEnumerable<ChangeBatchNoViewModel>> GetMaterialBatchesAsync(string materialId)
        {
            return await _changeBatchNoDataAccess.GetMaterialBatchesAsync(materialId);
        }

        public async Task SaveBatchModificationsAsync(IEnumerable<ChangeBatchNoViewModel> modifiedItems)
        {
            await _changeBatchNoDataAccess.SaveBatchModificationsAsync(modifiedItems);
        }
    }
}
