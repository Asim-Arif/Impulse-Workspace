using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public class MakerPOListService : IMakerPOListService
    {
        private readonly IMakerPOListDataAccess _dataAccess;

        public MakerPOListService(IMakerPOListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<(List<MakerPOListItem> Items, string ReportSql)> GetListAsync(MakerPOListFilter filter) => _dataAccess.GetListAsync(filter);
        public Task<List<LookupItemInt>> GetMakersAsync() => _dataAccess.GetMakersAsync();
        public Task<List<LookupItemString>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        public Task<List<LookupItemString>> GetItemCategoriesAsync() => _dataAccess.GetItemCategoriesAsync();
        public Task<List<LookupItemInt>> GetItemGroupsAsync() => _dataAccess.GetItemGroupsAsync();
        public Task<List<LookupItemString>> GetItemsAsync() => _dataAccess.GetItemsAsync();
        public Task<List<LookupItemInt>> GetProcessesAsync() => _dataAccess.GetProcessesAsync();
        public Task<List<LookupItemString>> GetEmployeesAsync() => _dataAccess.GetEmployeesAsync();
        public Task<bool> AuthorizeIssuancesAsync(IEnumerable<long> entryIds, string userName, string machineName) => _dataAccess.AuthorizeIssuancesAsync(entryIds, userName, machineName);
        public Task<bool> CloseMakerPOAsync(long entryId) => _dataAccess.CloseMakerPOAsync(entryId);
        public Task<(bool ShortLoan, bool LongLoan)> CheckLoanExistsAsync(string masterPoNo) => _dataAccess.CheckLoanExistsAsync(masterPoNo);
        public Task<int> CheckReceivingExistsAsync(long entryId) => _dataAccess.CheckReceivingExistsAsync(entryId);
        public Task<bool> DeleteIssuanceAsync(long entryId) => _dataAccess.DeleteIssuanceAsync(entryId);
        public Task<bool> GetUserRightAsync(string rightName, string userName) => _dataAccess.GetUserRightAsync(rightName, userName);
    }
}
