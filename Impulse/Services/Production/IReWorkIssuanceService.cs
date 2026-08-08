using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IReWorkIssuanceService
    {
        Task<List<LookupItemString>> GetArticlesAsync();
        Task<List<LookupItemInt>> GetRepairTypesAsync();
        Task<List<ReWorkAvailableItem>> GetAvailableReWorkLotsAsync(ReWorkIssuanceFilter filter);
        Task<List<LookupItemInt>> GetMakersForRepairProcessAsync(int repairRefId);
        Task<ReWorkIssuanceResult> SaveReWorkIssuanceAsync(SaveReWorkIssuanceRequest request);
    }
}
