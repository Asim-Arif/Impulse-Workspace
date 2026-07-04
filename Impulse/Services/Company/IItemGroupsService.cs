using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public interface IItemGroupsService
    {
        Task<List<ItemGroupsViewModel>> GetItemGroupsAsync();
        Task<List<PurchaserLookupModel>> GetPurchasersAsync();
        Task<bool> InsertItemGroupAsync(ItemGroupsViewModel model);
        Task<bool> UpdateItemGroupAsync(ItemGroupsViewModel model);
        Task<bool> DeleteItemGroupAsync(int id);
    }
}
