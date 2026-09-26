using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IPpcOrderPlanningService
    {
        Task<List<PpcOrderSummaryListItem>> GetAuthorizedOrdersForPlanningAsync();
        Task<PpcOrderHeaderDto?> GetOrderForPlanningAsync(string orderNo);
        Task<List<PurchaseProcessLookupDto>> GetPurchaseProcessesForItemAsync(string itemId);
        Task<List<MakerLookupDto>> GetMakersForItemAndProcessAsync(string itemId, int processId);
        Task<bool> SavePpcPlanAndDispatchTasksAsync(PpcPlanSaveRequest request);
    }
}
