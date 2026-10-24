using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IBiometricAgentCoordinator
    {
        bool HasConnectedAgent(string clientKey = "Default");
        void RegisterAgent(string connectionId, string clientKey, string locationName);
        void UnregisterAgent(string connectionId);
        void CompleteRequest(BiometricAgentReadResult result);
        Task<(bool success, string error, List<IFacePunchRecord> records)> RequestPunchesAsync(
            string ipAddress, int machineNo, int portNo = 4370, string prefix = "EMP", string clientKey = "Default", bool clearLogs = false);
    }
}
