using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IChangeBatchNoDataAccess
    {
        Task<IEnumerable<object>> GetMaterialsAsync();
        Task<IEnumerable<ChangeBatchNoViewModel>> GetMaterialBatchesAsync(string materialId);
        Task SaveBatchModificationsAsync(IEnumerable<ChangeBatchNoViewModel> modifiedItems);
    }
}
