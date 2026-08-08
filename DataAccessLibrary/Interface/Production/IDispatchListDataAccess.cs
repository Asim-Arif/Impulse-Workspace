using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IDispatchListDataAccess
    {
        Task<List<LookupItemString>> GetCustomersAsync();
        Task<List<DispatchListSummaryItem>> GetDispatchListsAsync(DispatchListSearchFilter filter);
        Task<FinalizeDispatchResult> FinalizeDispatchListAsync(long entryId, string userName, string machineName);
    }
}
