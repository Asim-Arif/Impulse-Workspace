using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IFinishStockIssuanceDataAccess
    {
        Task<string> GetNextIssuanceNoAsync(DateTime date);
        Task<List<CustomerLookupModel>> GetCustomersAsync();
        Task<List<ItemViewModel>> GetArticlesAsync(string custCode);
        Task<List<SFTargetStoreModel>> GetAvailableStoresAsync(string itemId);
        Task<List<LocationLotStockModel>> GetAvailableLocationsAsync(string itemId, int storeRefId, string custCode);
        Task<bool> SaveIssuanceAsync(FinishStockIssuanceMasterModel master, List<FinishStockIssuanceItemModel> items);
    }
}
