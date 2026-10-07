using System;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.IntraOffice;
using DataAccessLibrary.Models.IntraOffice;
using Microsoft.AspNetCore.SignalR;

namespace Impulse.Hubs
{
    public class ChatHub : Hub
    {
        private readonly IIntraOfficeDataAccess _dataAccess;

        public ChatHub(IIntraOfficeDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public override async Task OnConnectedAsync()
        {
            var userId = Context.User?.Identity?.Name 
                ?? Context.User?.FindFirstValue(ClaimTypes.Name) 
                ?? Context.User?.FindFirstValue(ClaimTypes.NameIdentifier);

            if (!string.IsNullOrEmpty(userId))
            {
                await _dataAccess.UpdateUserPresenceAsync(userId, UserStatus.Online, Context.ConnectionId);
                await Clients.Others.SendAsync("UserOnline", userId);
            }
            await base.OnConnectedAsync();
        }

        public override async Task OnDisconnectedAsync(Exception? exception)
        {
            var userId = Context.User?.Identity?.Name 
                ?? Context.User?.FindFirstValue(ClaimTypes.Name) 
                ?? Context.User?.FindFirstValue(ClaimTypes.NameIdentifier);

            if (!string.IsNullOrEmpty(userId))
            {
                await _dataAccess.UpdateUserPresenceAsync(userId, UserStatus.Offline);
                await Clients.Others.SendAsync("UserOffline", userId);
            }
            await base.OnDisconnectedAsync(exception);
        }

        public async Task SendMessageToChannel(int channelId, string senderName, string message)
        {
            var userId = Context.User?.Identity?.Name ?? senderName;
            await Clients.Group($"channel_{channelId}").SendAsync("ReceiveChannelMessage", channelId, userId, senderName, message, DateTime.UtcNow);
        }

        public async Task SendDirectMessage(string receiverId, string senderName, string message)
        {
            var senderId = Context.User?.Identity?.Name ?? senderName;
            await Clients.User(receiverId).SendAsync("ReceiveDirectMessage", senderId, senderName, message, DateTime.UtcNow);
            await Clients.Caller.SendAsync("ReceiveDirectMessage", senderId, senderName, message, DateTime.UtcNow);
        }

        public async Task JoinChannel(int channelId)
        {
            await Groups.AddToGroupAsync(Context.ConnectionId, $"channel_{channelId}");
        }

        public async Task LeaveChannel(int channelId)
        {
            await Groups.RemoveFromGroupAsync(Context.ConnectionId, $"channel_{channelId}");
        }

        public async Task SendTypingIndicator(int channelId, string userName)
        {
            await Clients.OthersInGroup($"channel_{channelId}").SendAsync("UserTyping", channelId, userName);
        }

        public async Task SendDirectTypingIndicator(string receiverId, string userName)
        {
            await Clients.User(receiverId).SendAsync("UserTypingDM", userName);
        }

        public async Task NotifyNewAnnouncement(string title)
        {
            await Clients.All.SendAsync("NewAnnouncement", title);
        }

        public async Task NotifyTaskUpdate(string assigneeId, string taskTitle, string status)
        {
            await Clients.User(assigneeId).SendAsync("TaskUpdated", taskTitle, status);
        }
    }
}
