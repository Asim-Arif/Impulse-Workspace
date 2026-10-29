using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IDispatchListService
    {
        Task<List<LookupItemString>> GetCustomersAsync();
        Task<List<DispatchListSummaryItem>> GetDispatchListsAsync(DispatchListSearchFilter filter);
        Task<FinalizeDispatchResult> FinalizeDispatchListAsync(long entryId, string userName, string machineName);
        Task<DeleteDispatchResult> DeleteDispatchListAsync(long entryId, string dispatchListNo, string userName, string machineName);
    }
}
