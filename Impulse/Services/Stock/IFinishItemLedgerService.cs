using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IFinishItemLedgerService
    {
        Task<ItemHeaderInfoModel?> GetItemHeaderInfoAsync(string itemId);
        Task<List<FinishItemLedgerRowModel>> GetLedgerAsync(FinishItemLedgerFilterModel filter);
        Task<List<ItemViewModel>> GetArticlesAsync();
    }
}
