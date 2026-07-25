using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Services.Stock
{
    public interface IChangeBatchNoService
    {
        Task<IEnumerable<object>> GetMaterialsAsync();
        Task<IEnumerable<ChangeBatchNoViewModel>> GetMaterialBatchesAsync(string materialId);
        Task SaveBatchModificationsAsync(IEnumerable<ChangeBatchNoViewModel> modifiedItems);
    }
}
