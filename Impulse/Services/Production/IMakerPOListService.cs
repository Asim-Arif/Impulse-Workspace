using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IMakerPOListService
    {
        Task<(List<MakerPOListItem> Items, string ReportSql)> GetListAsync(MakerPOListFilter filter);
        Task<List<LookupItemInt>> GetMakersAsync();
        Task<List<LookupItemString>> GetCustomersAsync();
        Task<List<LookupItemString>> GetItemCategoriesAsync();
        Task<List<LookupItemInt>> GetItemGroupsAsync();
        Task<List<LookupItemString>> GetItemsAsync();
        Task<List<LookupItemInt>> GetProcessesAsync();
        Task<List<LookupItemString>> GetEmployeesAsync();
        Task<bool> AuthorizeIssuancesAsync(IEnumerable<long> entryIds, string userName, string machineName);
        Task<bool> CloseMakerPOAsync(long entryId);
        Task<(bool ShortLoan, bool LongLoan)> CheckLoanExistsAsync(string masterPoNo);
        Task<int> CheckReceivingExistsAsync(long entryId);
        Task<bool> DeleteIssuanceAsync(long entryId);
        Task<bool> GetUserRightAsync(string rightName, string userName);
    }
}
