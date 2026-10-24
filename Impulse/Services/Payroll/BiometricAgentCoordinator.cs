using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Hubs;
using Microsoft.AspNetCore.SignalR;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Payroll
{
    public class BiometricAgentCoordinator : IBiometricAgentCoordinator
    {
        private readonly IHubContext<BiometricAgentHub> _hubContext;
        private readonly ILogger<BiometricAgentCoordinator> _logger;

        // Key = ConnectionId, Value = (ClientKey, LocationName)
        private readonly ConcurrentDictionary<string, (string ClientKey, string LocationName)> _connectedAgents = new();

        // Key = RequestId, Value = TaskCompletionSource
        private readonly ConcurrentDictionary<string, TaskCompletionSource<BiometricAgentReadResult>> _pendingRequests = new();

        public BiometricAgentCoordinator(
            IHubContext<BiometricAgentHub> hubContext,
            ILogger<BiometricAgentCoordinator> logger)
        {
            _hubContext = hubContext;
            _logger = logger;
        }

        public bool HasConnectedAgent(string clientKey = "Default")
        {
            return _connectedAgents.Values.Any(a => string.Equals(a.ClientKey, clientKey, StringComparison.OrdinalIgnoreCase));
        }

        public void RegisterAgent(string connectionId, string clientKey, string locationName)
        {
            _connectedAgents[connectionId] = (clientKey, locationName);
            _logger.LogInformation("Biometric Agent registered. ClientKey: {ClientKey}, Location: {Location}, ConnectionId: {ConnectionId}",
                clientKey, locationName, connectionId);
        }

        public void UnregisterAgent(string connectionId)
        {
            if (_connectedAgents.TryRemove(connectionId, out var agent))
            {
                _logger.LogInformation("Biometric Agent disconnected. ClientKey: {ClientKey}, ConnectionId: {ConnectionId}",
                    agent.ClientKey, connectionId);
            }
        }

        public void CompleteRequest(BiometricAgentReadResult result)
        {
            if (!string.IsNullOrEmpty(result.RequestId) &&
                _pendingRequests.TryGetValue(result.RequestId, out var tcs))
            {
                tcs.TrySetResult(result);
            }
        }

        public async Task<(bool success, string error, List<IFacePunchRecord> records)> RequestPunchesAsync(
            string ipAddress, int machineNo, int portNo = 4370, string prefix = "EMP", string clientKey = "Default", bool clearLogs = false)
        {
            // Find active agent for this clientKey
            var agentEntry = _connectedAgents.FirstOrDefault(a => string.Equals(a.Value.ClientKey, clientKey, StringComparison.OrdinalIgnoreCase));
            if (string.IsNullOrEmpty(agentEntry.Key))
            {
                return (false, $"No active sync agent found for client '{clientKey}'.", new());
            }

            string connectionId = agentEntry.Key;
            string requestId = Guid.NewGuid().ToString("N");
            var tcs = new TaskCompletionSource<BiometricAgentReadResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            _pendingRequests[requestId] = tcs;

            var cmd = new BiometricAgentReadCommand
            {
                RequestId = requestId,
                IPAddress = ipAddress,
                MachineNo = machineNo,
                PortNo = portNo,
                Prefix = prefix,
                ClearLogs = clearLogs
            };

            try
            {
                _logger.LogInformation("Dispatching ExecuteRead to agent {ConnectionId} for IP {IP}", connectionId, ipAddress);
                await _hubContext.Clients.Client(connectionId).SendAsync("ExecuteRead", cmd);

                // Wait up to 45 seconds for agent to connect, read, and return punches
                var timeoutTask = Task.Delay(TimeSpan.FromSeconds(45));
                var completedTask = await Task.WhenAny(tcs.Task, timeoutTask);

                if (completedTask == tcs.Task)
                {
                    var result = await tcs.Task;
                    return (result.Success, result.ErrorMessage, result.Punches ?? new());
                }
                else
                {
                    _logger.LogWarning("Remote agent read timed out for IP {IP}", ipAddress);
                    return (false, $"Sync Agent read timed out while connecting to machine ({ipAddress}).", new());
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to dispatch read command to agent {ConnectionId}", connectionId);
                return (false, $"Failed to dispatch command to sync agent: {ex.Message}", new());
            }
            finally
            {
                _pendingRequests.TryRemove(requestId, out _);
            }
        }
    }
}
