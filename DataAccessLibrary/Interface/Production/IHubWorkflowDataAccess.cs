using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IHubWorkflowDataAccess
    {
        Task<HubDetectionResultDto?> AnalyzeLotProcessHubAsync(string itemCode, int processId);
        Task<bool> IsProcessAuthRequiredAsync(int processId);
        Task<int> CloseHubTaskAsync(string lotNo, string hubName, string completedByUserName, string notes);
        Task<int> CreateNextHubTaskAndNotificationsAsync(HubDetectionResultDto transition, string lotNo, string itemCode, string orderNo, decimal rcvdQty, string createdBy);
        Task<int> SendDispatchFinalLotNotificationAsync(string lotNo, string itemCode, string orderNo, decimal rcvdQty, string completedProcessName, string createdBy);
    }
}
