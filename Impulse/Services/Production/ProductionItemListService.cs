using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class ProductionItemListService : IProductionItemListService
    {
        private readonly IProductionItemListDataAccess _dataAccess;

        public ProductionItemListService(IProductionItemListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ProductionItemListItem>> GetListAsync(ProductionItemListFilters filters)
            => _dataAccess.GetListAsync(filters);

        public async Task<ProductionItemListLookups> GetInitialLookupsAsync()
        {
            // Load all non-cascading lookups in parallel for performance
            var customersTask = _dataAccess.GetCustomersAsync();
            var itemGroupsTask = _dataAccess.GetItemGroupsAsync();
            var processesTask = _dataAccess.GetProcessesAsync();
            var makersTask = _dataAccess.GetMakersAsync();
            var itemsTask = _dataAccess.GetItemsAsync();
            var hubsTask = _dataAccess.GetHubsAsync();

            await Task.WhenAll(customersTask, itemGroupsTask, processesTask, makersTask, itemsTask, hubsTask);

            return new ProductionItemListLookups
            {
                Customers = customersTask.Result,
                ItemGroups = itemGroupsTask.Result,
                Processes = processesTask.Result,
                Makers = makersTask.Result,
                Items = itemsTask.Result,
                Hubs = hubsTask.Result
            };
        }

        public Task<List<LookupItemString>> GetCountriesByCustomerAsync(string custCode)
            => _dataAccess.GetCountriesByCustomerAsync(custCode);

        public Task<List<LookupItemString>> GetOrdersByCustomerAsync(string custCode)
            => _dataAccess.GetOrdersByCustomerAsync(custCode);

        public Task<bool> SaveRemarksAsync(string lotNo, int entryType, string orderNo, string itemCode, string remarks)
            => _dataAccess.SaveRemarksAsync(lotNo, entryType, orderNo, itemCode, remarks);
    }
}
