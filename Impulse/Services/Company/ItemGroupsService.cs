using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public class ItemGroupsService : IItemGroupsService
    {
        private readonly IItemGroupsDataAccess _dataAccess;

        public ItemGroupsService(IItemGroupsDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ItemGroupsViewModel>> GetItemGroupsAsync()
        {
            return _dataAccess.GetItemGroupsAsync();
        }

        public Task<List<PurchaserLookupModel>> GetPurchasersAsync()
        {
            return _dataAccess.GetPurchasersAsync();
        }

        public Task<bool> InsertItemGroupAsync(ItemGroupsViewModel model)
        {
            return _dataAccess.InsertItemGroupAsync(model);
        }

        public Task<bool> UpdateItemGroupAsync(ItemGroupsViewModel model)
        {
            return _dataAccess.UpdateItemGroupAsync(model);
        }

        public Task<bool> DeleteItemGroupAsync(int id)
        {
            return _dataAccess.DeleteItemGroupAsync(id);
        }
    }
}
