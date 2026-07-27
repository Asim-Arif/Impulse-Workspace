using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IFinishItemLedgerDataAccess
    {
        Task<ItemHeaderInfoModel?> GetItemHeaderInfoAsync(string itemId);
        Task<List<FinishItemLedgerRowModel>> GetLedgerAsync(FinishItemLedgerFilterModel filter);
        Task<List<ItemViewModel>> GetArticlesAsync();
    }
}
