using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IHubWorkflowOrchestrator
    {
        Task<bool> IsAuthRequiredAsync(int processId);

        Task<HubTransitionExecutionResult> HandleLotReceivingHubTransitionAsync(
            string lotNo,
            int processId,
            string itemCode,
            string orderNo,
            decimal rcvdQty,
            string actorUserName);
    }
}
