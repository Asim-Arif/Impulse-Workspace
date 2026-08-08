using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class CreateDispatchListService : ICreateDispatchListService
    {
        private readonly ICreateDispatchListDataAccess _dataAccess;

        public CreateDispatchListService(ICreateDispatchListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<LookupItemString>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        public Task<List<string>> GetCountriesByCustomerAsync(string custCode) => _dataAccess.GetCountriesByCustomerAsync(custCode);
        public Task<List<LookupItemInt>> GetMakersAsync() => _dataAccess.GetMakersAsync();
        public Task<List<LookupItemString>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();
        public Task<List<FinalizedLotItem>> GetAvailableFinalizedLotsAsync(DispatchListFilter filter) => _dataAccess.GetAvailableFinalizedLotsAsync(filter);
        public Task<string?> GetInnerLabelReportNameAsync(string custCode, string country, bool otherLabel = false, bool manualQty = false)
            => _dataAccess.GetInnerLabelReportNameAsync(custCode, country, otherLabel, manualQty);
        public Task<LoadedDispatchListForEdit?> GetDispatchListForEditAsync(long entryId) => _dataAccess.GetDispatchListForEditAsync(entryId);
        public Task<SaveDispatchListResult> SaveDispatchListAsync(SaveDispatchListRequest request) => _dataAccess.SaveDispatchListAsync(request);
    }
}
