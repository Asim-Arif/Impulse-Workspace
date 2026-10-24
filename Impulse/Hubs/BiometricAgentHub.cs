using System;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.SignalR;

namespace Impulse.Hubs
{
    public class BiometricAgentHub : Hub
    {
        private readonly IBiometricAgentCoordinator _coordinator;

        public BiometricAgentHub(IBiometricAgentCoordinator coordinator)
        {
            _coordinator = coordinator;
        }

        public async Task RegisterAgent(string clientKey, string locationName)
        {
            _coordinator.RegisterAgent(Context.ConnectionId, clientKey, locationName);
            await Groups.AddToGroupAsync(Context.ConnectionId, clientKey);
        }

        public Task ReportReadResult(BiometricAgentReadResult result)
        {
            _coordinator.CompleteRequest(result);
            return Task.CompletedTask;
        }

        public override Task OnDisconnectedAsync(Exception? exception)
        {
            _coordinator.UnregisterAgent(Context.ConnectionId);
            return base.OnDisconnectedAsync(exception);
        }
    }
}
