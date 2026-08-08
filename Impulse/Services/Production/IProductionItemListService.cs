using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IProductionItemListService
    {
        Task<List<ProductionItemListItem>> GetListAsync(ProductionItemListFilters filters);
        Task<ProductionItemListLookups> GetInitialLookupsAsync();
        Task<List<LookupItemString>> GetCountriesByCustomerAsync(string custCode);
        Task<List<LookupItemString>> GetOrdersByCustomerAsync(string custCode);
        Task<bool> SaveRemarksAsync(string lotNo, int entryType, string orderNo, string itemCode, string remarks);
    }
}
