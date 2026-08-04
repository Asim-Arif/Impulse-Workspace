using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerRcvListService : IMakerRcvListService
    {
        private readonly IMakerRcvListDataAccess _dataAccess;

        public MakerRcvListService(IMakerRcvListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<(List<MakerRcvListItem> Items, string ReportSql)> GetListAsync(MakerRcvListFilter filter) => _dataAccess.GetListAsync(filter);
        public Task<List<LookupItemInt>> GetMakersAsync() => _dataAccess.GetMakersAsync();
        public Task<List<LookupItemString>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        public Task<List<LookupItemString>> GetItemCategoriesAsync() => _dataAccess.GetItemCategoriesAsync();
        public Task<List<LookupItemInt>> GetItemGroupsAsync() => _dataAccess.GetItemGroupsAsync();
        public Task<List<LookupItemString>> GetItemsAsync() => _dataAccess.GetItemsAsync();
        public Task<List<LookupItemInt>> GetProcessesAsync() => _dataAccess.GetProcessesAsync();
        public Task<bool> DeleteReceivingAsync(long vrdEntryId, string userName, string machineName) => _dataAccess.DeleteReceivingAsync(vrdEntryId, userName, machineName);
        public Task<bool> CloseLotAsync(string lotNo, string userName, string machineName) => _dataAccess.CloseLotAsync(lotNo, userName, machineName);
        public Task<bool> UpdateManualPTCNoAsync(string lotNo, string manualPTCNo) => _dataAccess.UpdateManualPTCNoAsync(lotNo, manualPTCNo);
        public Task<int> CheckIssuanceExistsAsync(long vrdEntryId) => _dataAccess.CheckIssuanceExistsAsync(vrdEntryId);
        public Task<bool> GetUserRightAsync(string rightName, string userName) => _dataAccess.GetUserRightAsync(rightName, userName);
    }
}
