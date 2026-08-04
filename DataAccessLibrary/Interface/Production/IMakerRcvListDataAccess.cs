using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IMakerRcvListDataAccess
    {
        Task<(List<MakerRcvListItem> Items, string ReportSql)> GetListAsync(MakerRcvListFilter filter);
        Task<List<LookupItemInt>> GetMakersAsync();
        Task<List<LookupItemString>> GetCustomersAsync();
        Task<List<LookupItemString>> GetItemCategoriesAsync();
        Task<List<LookupItemInt>> GetItemGroupsAsync();
        Task<List<LookupItemString>> GetItemsAsync();
        Task<List<LookupItemInt>> GetProcessesAsync();
        Task<bool> DeleteReceivingAsync(long vrdEntryId, string userName, string machineName);
        Task<bool> CloseLotAsync(string lotNo, string userName, string machineName);
        Task<bool> UpdateManualPTCNoAsync(string lotNo, string manualPTCNo);
        Task<int> CheckIssuanceExistsAsync(long vrdEntryId);
        Task<bool> GetUserRightAsync(string rightName, string userName);
    }
}
