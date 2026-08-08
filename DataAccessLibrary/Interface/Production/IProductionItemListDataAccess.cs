using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IProductionItemListDataAccess
    {
        Task<List<ProductionItemListItem>> GetListAsync(ProductionItemListFilters filters);
        Task<List<LookupItemString>> GetCustomersAsync();
        Task<List<LookupItemString>> GetCountriesByCustomerAsync(string custCode);
        Task<List<LookupItemString>> GetOrdersByCustomerAsync(string custCode);
        Task<List<LookupItemInt>> GetItemGroupsAsync();
        Task<List<LookupItemInt>> GetProcessesAsync();
        Task<List<LookupItemInt>> GetMakersAsync();
        Task<List<LookupItemString>> GetItemsAsync();
        Task<List<string>> GetHubsAsync();
        Task<bool> SaveRemarksAsync(string lotNo, int entryType, string orderNo, string itemCode, string remarks);
    }
}
