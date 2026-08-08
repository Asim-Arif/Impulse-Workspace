using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class ReWorkIssuanceService : IReWorkIssuanceService
    {
        private readonly IReWorkIssuanceDataAccess _dataAccess;

        public ReWorkIssuanceService(IReWorkIssuanceDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<LookupItemString>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();
        public Task<List<LookupItemInt>> GetRepairTypesAsync() => _dataAccess.GetRepairTypesAsync();
        public Task<List<ReWorkAvailableItem>> GetAvailableReWorkLotsAsync(ReWorkIssuanceFilter filter) => _dataAccess.GetAvailableReWorkLotsAsync(filter);
        public Task<List<LookupItemInt>> GetMakersForRepairProcessAsync(int repairRefId) => _dataAccess.GetMakersForRepairProcessAsync(repairRefId);
        public Task<ReWorkIssuanceResult> SaveReWorkIssuanceAsync(SaveReWorkIssuanceRequest request) => _dataAccess.SaveReWorkIssuanceAsync(request);
    }
}
