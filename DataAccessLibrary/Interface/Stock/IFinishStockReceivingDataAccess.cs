using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IFinishStockReceivingDataAccess
    {
        Task<string> GetNextReceivingNoAsync(DateTime date);
        Task<string> GetNextLotNoAsync();
        Task<List<CustomerLookupModel>> GetCustomersAsync();
        Task<List<ItemViewModel>> GetArticlesAsync(string custCode);
        Task<List<SFTargetStoreModel>> GetStoresAsync(string userName);
        Task<List<SFTargetShelfModel>> GetShelvesAsync(int storeRefId);
        Task<bool> IsLotNoExistsAsync(string lotNo);
        Task<bool> IsBatchNoExistsAsync(string batchNo);
        Task<bool> SaveReceivingAsync(FinishStockReceivingMasterModel master, List<FinishStockReceivingItemModel> items);
    }
}
