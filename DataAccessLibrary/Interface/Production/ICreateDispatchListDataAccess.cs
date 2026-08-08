using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface ICreateDispatchListDataAccess
    {
        Task<List<LookupItemString>> GetCustomersAsync();
        Task<List<string>> GetCountriesByCustomerAsync(string custCode);
        Task<List<LookupItemInt>> GetMakersAsync();
        Task<List<LookupItemString>> GetArticlesAsync();
        Task<List<FinalizedLotItem>> GetAvailableFinalizedLotsAsync(DispatchListFilter filter);
        Task<string?> GetInnerLabelReportNameAsync(string custCode, string country, bool otherLabel = false, bool manualQty = false);
        Task<LoadedDispatchListForEdit?> GetDispatchListForEditAsync(long entryId);
        Task<SaveDispatchListResult> SaveDispatchListAsync(SaveDispatchListRequest request);
    }
}
