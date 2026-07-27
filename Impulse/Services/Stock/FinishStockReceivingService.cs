using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class FinishStockReceivingService : IFinishStockReceivingService
    {
        private readonly IFinishStockReceivingDataAccess _dataAccess;

        public FinishStockReceivingService(IFinishStockReceivingDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<string> GetNextReceivingNoAsync(DateTime date) => _dataAccess.GetNextReceivingNoAsync(date);

        public Task<string> GetNextLotNoAsync() => _dataAccess.GetNextLotNoAsync();

        public Task<List<CustomerLookupModel>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();

        public Task<List<ItemViewModel>> GetArticlesAsync(string custCode) => _dataAccess.GetArticlesAsync(custCode);

        public Task<List<SFTargetStoreModel>> GetStoresAsync(string userName) => _dataAccess.GetStoresAsync(userName);

        public Task<List<SFTargetShelfModel>> GetShelvesAsync(int storeRefId) => _dataAccess.GetShelvesAsync(storeRefId);

        public Task<bool> IsLotNoExistsAsync(string lotNo) => _dataAccess.IsLotNoExistsAsync(lotNo);

        public Task<bool> IsBatchNoExistsAsync(string batchNo) => _dataAccess.IsBatchNoExistsAsync(batchNo);

        public Task<bool> SaveReceivingAsync(FinishStockReceivingMasterModel master, List<FinishStockReceivingItemModel> items) => _dataAccess.SaveReceivingAsync(master, items);
    }
}
