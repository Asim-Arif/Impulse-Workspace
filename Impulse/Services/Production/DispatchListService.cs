using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class DispatchListService : IDispatchListService
    {
        private readonly IDispatchListDataAccess _dataAccess;

        public DispatchListService(IDispatchListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<LookupItemString>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        public Task<List<DispatchListSummaryItem>> GetDispatchListsAsync(DispatchListSearchFilter filter) => _dataAccess.GetDispatchListsAsync(filter);
        public Task<FinalizeDispatchResult> FinalizeDispatchListAsync(long entryId, string userName, string machineName) => _dataAccess.FinalizeDispatchListAsync(entryId, userName, machineName);
    }
}
