using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class FinishItemLedgerService : IFinishItemLedgerService
    {
        private readonly IFinishItemLedgerDataAccess _dataAccess;

        public FinishItemLedgerService(IFinishItemLedgerDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<ItemHeaderInfoModel?> GetItemHeaderInfoAsync(string itemId) => _dataAccess.GetItemHeaderInfoAsync(itemId);

        public Task<List<FinishItemLedgerRowModel>> GetLedgerAsync(FinishItemLedgerFilterModel filter) => _dataAccess.GetLedgerAsync(filter);

        public Task<List<ItemViewModel>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();
    }
}
