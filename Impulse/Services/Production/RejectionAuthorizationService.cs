using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public class RejectionAuthorizationService : IRejectionAuthorizationService
    {
        private readonly IRejectionAuthorizationDataAccess _dataAccess;

        public RejectionAuthorizationService(IRejectionAuthorizationDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<int> GetFactoryMakerIdAsync() => _dataAccess.GetFactoryMakerIdAsync();

        public Task<List<RejectionAuthMakerLookup>> GetMakersAsync(int factoryMakerId) => _dataAccess.GetMakersAsync(factoryMakerId);

        public Task<List<RejectionAuthEmployeeLookup>> GetEmployeesAsync() => _dataAccess.GetEmployeesAsync();

        public Task<List<RejectionAuthCategoryLookup>> GetCategoriesAsync() => _dataAccess.GetCategoriesAsync();

        public Task<List<RejectionAuthItemModel>> GetPendingRejectionsAsync(RejectionAuthFilterModel filter) => _dataAccess.GetPendingRejectionsAsync(filter);

        public Task<List<RejectionAuthItemModel>> GetHistoryAsync(RejectionAuthFilterModel filter) => _dataAccess.GetHistoryAsync(filter);

        public Task<int> UpdateDeductFromBillingAsync(List<int> entryIds, bool deduct) => _dataAccess.UpdateDeductFromBillingAsync(entryIds, deduct);

        public Task<bool> UpdateRateAsync(int entryId, int rate) => _dataAccess.UpdateRateAsync(entryId, rate);
    }
}
