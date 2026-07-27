using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class FinishStockIssuanceService : IFinishStockIssuanceService
    {
        private readonly IFinishStockIssuanceDataAccess _dataAccess;

        public FinishStockIssuanceService(IFinishStockIssuanceDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<string> GetNextIssuanceNoAsync(DateTime date) => _dataAccess.GetNextIssuanceNoAsync(date);

        public Task<List<CustomerLookupModel>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();

        public Task<List<ItemViewModel>> GetArticlesAsync(string custCode) => _dataAccess.GetArticlesAsync(custCode);

        public Task<List<SFTargetStoreModel>> GetAvailableStoresAsync(string itemId) => _dataAccess.GetAvailableStoresAsync(itemId);

        public Task<List<LocationLotStockModel>> GetAvailableLocationsAsync(string itemId, int storeRefId, string custCode) => _dataAccess.GetAvailableLocationsAsync(itemId, storeRefId, custCode);

        public Task<bool> SaveIssuanceAsync(FinishStockIssuanceMasterModel master, List<FinishStockIssuanceItemModel> items) => _dataAccess.SaveIssuanceAsync(master, items);
    }
}
