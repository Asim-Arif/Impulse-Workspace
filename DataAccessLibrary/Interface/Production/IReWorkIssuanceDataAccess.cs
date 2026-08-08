using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IReWorkIssuanceDataAccess
    {
        Task<List<LookupItemString>> GetArticlesAsync();
        Task<List<LookupItemInt>> GetRepairTypesAsync();
        Task<List<ReWorkAvailableItem>> GetAvailableReWorkLotsAsync(ReWorkIssuanceFilter filter);
        Task<List<LookupItemInt>> GetMakersForRepairProcessAsync(int repairRefId);
        Task<ReWorkIssuanceResult> SaveReWorkIssuanceAsync(SaveReWorkIssuanceRequest request);
    }
}
