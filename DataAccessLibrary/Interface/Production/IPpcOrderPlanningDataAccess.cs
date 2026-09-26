using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IPpcOrderPlanningDataAccess
    {
        Task<List<PpcOrderSummaryListItem>> GetAuthorizedOrdersForPlanningAsync();
        Task<PpcOrderHeaderDto?> GetOrderForPlanningAsync(string orderNo);
        Task<List<PurchaseProcessLookupDto>> GetPurchaseProcessesForItemAsync(string itemId);
        Task<List<MakerLookupDto>> GetMakersForItemAndProcessAsync(string itemId, int processId);
        Task<bool> SavePpcPlanAsync(PpcPlanSaveRequest request);
    }
}
