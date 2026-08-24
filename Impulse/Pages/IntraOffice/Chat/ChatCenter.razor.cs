using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.SignalR.Client;
using Radzen;

namespace Impulse.Pages.IntraOffice.Chat
{
    public partial class ChatCenter : ComponentBase
    {
        [Parameter] public string? OtherUserId { get; set; }
        [Parameter] public int? ChannelId { get; set; }

        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private IAiAssistantService AiService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;

        protected string CurrentUserId { get; set; } = string.Empty;
        protected string ActiveTab { get; set; } = "direct";
        protected string SearchQuery { get; set; } = string.Empty;

        protected bool IsLoadingSidebar { get; set; } = true;
        protected bool IsMessagesLoading { get; set; } = false;
        protected bool IsSending { get; set; } = false;
        protected bool IsBotActive { get; set; } = false;

        protected string InputMessageText { get; set; } = string.Empty;
        protected string TypingIndicatorText { get; set; } = string.Empty;

        protected List<IntraUserProfile> UsersList { get; set; } = new();
        protected List<Channel> ChannelsList { get; set; } = new();
        protected List<(string DeptId, string DeptName)> Departments { get; set; } = new();
        protected Dictionary<string, int> UnreadCounts { get; set; } = new();

        protected List<Message> CurrentMessages { get; set; } = new();
        protected Channel? CurrentChannel { get; set; }
        protected IntraUserProfile? CurrentOtherUser { get; set; }

        // New Channel Form
        protected bool ShowCreateChannelModalState { get; set; } = false;
        protected string NewChannelName { get; set; } = string.Empty;
        protected string NewChannelDescription { get; set; } = string.Empty;
        protected string NewChannelDepartmentId { get; set; } = string.Empty;

        private HubConnection? _hubConnection;

        protected bool ActiveConversationSelected => IsBotActive || !string.IsNullOrEmpty(OtherUserId) || ChannelId.HasValue;

        protected string ActiveTargetTitle
        {
            get
            {
                if (IsBotActive) return "IntraBot AI Assistant";
                if (ChannelId.HasValue) return CurrentChannel?.Name ?? $"Channel #{ChannelId}";
                return CurrentOtherUser?.EmployeeName ?? CurrentOtherUser?.FullUserName ?? OtherUserId ?? "";
            }
        }

        protected string ActiveTargetSubtitle
        {
            get
            {
                if (IsBotActive) return "Always Online · Impulse AI Assistant";
                if (ChannelId.HasValue) return $"{CurrentChannel?.DepartmentName ?? "Public"} · {CurrentChannel?.MemberCount ?? 0} members";
                return $"{CurrentOtherUser?.Designation ?? "Colleague"} · {CurrentOtherUser?.DepartmentName ?? "Department"}";
            }
        }

        protected List<IntraUserProfile> FilteredUsers
        {
            get
            {
                var query = UsersList.Where(u => !string.Equals(u.UserName, CurrentUserId, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrWhiteSpace(SearchQuery))
                {
                    query = query.Where(u =>
                        u.UserName.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase) ||
                        (u.FullUserName?.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        (u.EmployeeName?.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        (u.Designation?.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase) ?? false));
                }
                return query.ToList();
            }
        }

        protected List<Channel> FilteredChannels
        {
            get
            {
                var query = ChannelsList.AsEnumerable();
                if (!string.IsNullOrWhiteSpace(SearchQuery))
                {
                    query = query.Where(c => c.Name.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase));
                }
                return query.ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;
            CurrentUserId = user.Identity?.Name ?? user.FindFirst(ClaimTypes.Name)?.Value ?? "Guest";

            await LoadSidebarDataAsync();
            _ = InitializeSignalRAsync();
        }

        protected override async Task OnParametersSetAsync()
        {
            if (!string.IsNullOrEmpty(OtherUserId))
            {
                ActiveTab = "direct";
                await SelectUser(OtherUserId, updateUrl: false);
            }
            else if (ChannelId.HasValue)
            {
                ActiveTab = "channels";
                await SelectChannel(ChannelId.Value, updateUrl: false);
            }
        }

        protected async Task LoadSidebarDataAsync()
        {
            IsLoadingSidebar = true;
            try
            {
                UsersList = await IntraOfficeService.GetActiveUsersAsync();
                ChannelsList = await IntraOfficeService.GetChannelsForUserAsync(CurrentUserId);
                Departments = await IntraOfficeService.GetDepartmentsAsync();

                var directSummaries = await IntraOfficeService.GetDirectChatSummariesAsync(CurrentUserId);
                UnreadCounts = directSummaries.ToDictionary(s => s.OtherUserName, s => s.UnreadCount);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Chats",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoadingSidebar = false;
                StateHasChanged();
            }
        }

        private async Task InitializeSignalRAsync()
        {
            try
            {
                _hubConnection = new HubConnectionBuilder()
                    .WithUrl(NavigationManager.ToAbsoluteUri("/chathub"))
                    .WithAutomaticReconnect()
                    .Build();

                _hubConnection.On<long, string, string, string, DateTime>("ReceiveDirectMessage", async (id, senderId, senderName, content, sentAt) =>
                {
                    if (string.Equals(OtherUserId, senderId, StringComparison.OrdinalIgnoreCase) || string.Equals(CurrentUserId, senderId, StringComparison.OrdinalIgnoreCase))
                    {
                        CurrentMessages.Add(new Message
                        {
                            Id = id,
                            SenderId = senderId,
                            SenderName = senderName,
                            Content = content,
                            SentAt = sentAt
                        });
                        await InvokeAsync(StateHasChanged);
                    }
                    else
                    {
                        if (UnreadCounts.ContainsKey(senderId)) UnreadCounts[senderId]++;
                        else UnreadCounts[senderId] = 1;
                        await InvokeAsync(StateHasChanged);
                    }
                });

                _hubConnection.On<int, string, string, string, DateTime>("ReceiveChannelMessage", async (channelId, senderId, senderName, content, sentAt) =>
                {
                    if (ChannelId == channelId)
                    {
                        CurrentMessages.Add(new Message
                        {
                            ChannelId = channelId,
                            SenderId = senderId,
                            SenderName = senderName,
                            Content = content,
                            SentAt = sentAt
                        });
                        await InvokeAsync(StateHasChanged);
                    }
                });

                _hubConnection.On<string>("UserTypingDM", (userName) =>
                {
                    TypingIndicatorText = $"{userName} is typing...";
                    InvokeAsync(StateHasChanged);
                    _ = ClearTypingAfterDelay();
                });

                await _hubConnection.StartAsync();
            }
            catch (Exception)
            {
                // Fallback graceful
            }
        }

        private async Task ClearTypingAfterDelay()
        {
            await Task.Delay(3000);
            TypingIndicatorText = string.Empty;
            await InvokeAsync(StateHasChanged);
        }

        protected void SetActiveTab(string tab)
        {
            ActiveTab = tab;
        }

        protected async Task SelectUser(string userName, bool updateUrl = true)
        {
            IsBotActive = false;
            ChannelId = null;
            OtherUserId = userName;
            CurrentOtherUser = UsersList.FirstOrDefault(u => string.Equals(u.UserName, userName, StringComparison.OrdinalIgnoreCase));

            if (UnreadCounts.ContainsKey(userName))
            {
                UnreadCounts[userName] = 0;
            }

            IsMessagesLoading = true;
            StateHasChanged();

            try
            {
                CurrentMessages = await IntraOfficeService.GetDirectMessagesAsync(CurrentUserId, userName);
                await IntraOfficeService.MarkMessagesAsReadAsync(CurrentUserId, userName);
            }
            finally
            {
                IsMessagesLoading = false;
                StateHasChanged();
            }
        }

        protected async Task SelectChannel(int channelId, bool updateUrl = true)
        {
            IsBotActive = false;
            OtherUserId = null;
            ChannelId = channelId;
            CurrentChannel = await IntraOfficeService.GetChannelByIdAsync(channelId);

            if (CurrentChannel != null) CurrentChannel.UnreadCount = 0;

            if (_hubConnection?.State == HubConnectionState.Connected)
            {
                await _hubConnection.SendAsync("JoinChannel", channelId);
            }

            IsMessagesLoading = true;
            StateHasChanged();

            try
            {
                CurrentMessages = await IntraOfficeService.GetChannelMessagesAsync(channelId);
                await IntraOfficeService.MarkMessagesAsReadAsync(CurrentUserId, channelId: channelId);
            }
            finally
            {
                IsMessagesLoading = false;
                StateHasChanged();
            }
        }

        protected void SelectBot()
        {
            IsBotActive = true;
            OtherUserId = null;
            ChannelId = null;
            CurrentMessages = new List<Message>
            {
                new()
                {
                    SenderId = "IntraBot",
                    SenderName = "IntraBot AI",
                    Content = "Hello! I am IntraBot, your smart assistant for Impulse ERP. How can I help you today?",
                    SentAt = DateTime.UtcNow
                }
            };
        }

        protected async Task HandleInputKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter" && !string.IsNullOrWhiteSpace(InputMessageText) && !IsSending)
            {
                await SendMessageAsync();
            }
        }

        protected async Task SendMessageAsync()
        {
            if (string.IsNullOrWhiteSpace(InputMessageText)) return;

            var text = InputMessageText.Trim();
            InputMessageText = string.Empty;
            IsSending = true;

            try
            {
                if (IsBotActive)
                {
                    CurrentMessages.Add(new Message
                    {
                        SenderId = CurrentUserId,
                        SenderName = "You",
                        Content = text,
                        SentAt = DateTime.UtcNow
                    });

                    var botReply = await AiService.GenerateResponseAsync(text, CurrentUserId);
                    CurrentMessages.Add(new Message
                    {
                        SenderId = "IntraBot",
                        SenderName = "IntraBot AI",
                        Content = botReply,
                        SentAt = DateTime.UtcNow
                    });
                }
                else if (ChannelId.HasValue)
                {
                    var msg = new Message
                    {
                        ChannelId = ChannelId.Value,
                        SenderId = CurrentUserId,
                        Content = text,
                        MessageType = MessageType.Text
                    };

                    var id = await IntraOfficeService.SendMessageAsync(msg);
                    msg.Id = id;
                    msg.SenderName = "You";
                    CurrentMessages.Add(msg);

                    if (_hubConnection?.State == HubConnectionState.Connected)
                    {
                        await _hubConnection.SendAsync("SendMessageToChannel", ChannelId.Value, CurrentUserId, text);
                    }
                }
                else if (!string.IsNullOrEmpty(OtherUserId))
                {
                    var msg = new Message
                    {
                        ReceiverId = OtherUserId,
                        SenderId = CurrentUserId,
                        Content = text,
                        MessageType = MessageType.Text
                    };

                    var id = await IntraOfficeService.SendMessageAsync(msg);
                    msg.Id = id;
                    msg.SenderName = "You";
                    CurrentMessages.Add(msg);

                    if (_hubConnection?.State == HubConnectionState.Connected)
                    {
                        await _hubConnection.SendAsync("SendDirectMessage", OtherUserId, CurrentUserId, text);
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Send Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSending = false;
                StateHasChanged();
            }
        }

        protected void StartMeetingWithUser()
        {
            if (string.IsNullOrEmpty(OtherUserId)) return;
            NavigationManager.NavigateTo($"/intraoffice/meetings?withUser={OtherUserId}");
        }

        protected void OpenCreateChannelModal()
        {
            NewChannelName = string.Empty;
            NewChannelDescription = string.Empty;
            NewChannelDepartmentId = string.Empty;
            ShowCreateChannelModalState = true;
        }

        protected void CloseCreateChannelModal()
        {
            ShowCreateChannelModalState = false;
        }

        protected async Task SaveChannelAsync()
        {
            if (string.IsNullOrWhiteSpace(NewChannelName))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter a channel name.",
                    Duration = 3000
                });
                return;
            }

            try
            {
                var ch = new Channel
                {
                    Name = NewChannelName.Trim(),
                    Description = string.IsNullOrWhiteSpace(NewChannelDescription) ? null : NewChannelDescription.Trim(),
                    DepartmentId = string.IsNullOrWhiteSpace(NewChannelDepartmentId) ? null : NewChannelDepartmentId,
                    CreatedBy = CurrentUserId,
                    IsActive = true
                };

                var id = await IntraOfficeService.CreateChannelAsync(ch);
                if (id > 0)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Created",
                        Detail = "Channel created successfully.",
                        Duration = 3000
                    });

                    ShowCreateChannelModalState = false;
                    await LoadSidebarDataAsync();
                    await SelectChannel(id);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (_hubConnection != null)
            {
                await _hubConnection.DisposeAsync();
            }
        }
    }
}
