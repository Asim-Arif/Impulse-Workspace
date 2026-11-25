using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IRejectionAuthorizationDataAccess
    {
        Task<int> GetFactoryMakerIdAsync();
        Task<List<RejectionAuthMakerLookup>> GetMakersAsync(int factoryMakerId);
        Task<List<RejectionAuthEmployeeLookup>> GetEmployeesAsync();
        Task<List<RejectionAuthCategoryLookup>> GetCategoriesAsync();
        Task<List<RejectionAuthItemModel>> GetPendingRejectionsAsync(RejectionAuthFilterModel filter);
        Task<List<RejectionAuthItemModel>> GetHistoryAsync(RejectionAuthFilterModel filter);
        Task<int> UpdateDeductFromBillingAsync(List<int> entryIds, bool deduct);
        Task<bool> UpdateRateAsync(int entryId, int rate);
    }
}
