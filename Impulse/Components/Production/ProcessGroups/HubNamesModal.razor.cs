using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Components.Production.ProcessGroups
{
    public partial class HubNamesModal : ComponentBase
    {
        [Inject] public IProcessGroupService ProcessGroupService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public bool IsOpen { get; set; }
        [Parameter] public EventCallback OnClosed { get; set; }
        [Parameter] public EventCallback OnChanged { get; set; }

        public List<string> HubNames { get; set; } = new();
        public string NewHubName { get; set; } = string.Empty;
        public bool IsLoading { get; set; } = false;
        public bool IsProcessing { get; set; } = false;

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen)
            {
                await LoadHubNamesAsync();
            }
        }

        public async Task LoadHubNamesAsync()
        {
            try
            {
                IsLoading = true;
                HubNames = await ProcessGroupService.GetHubNamesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load hub names: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task AddHub()
        {
            if (string.IsNullOrWhiteSpace(NewHubName)) return;

            try
            {
                IsProcessing = true;
                bool success = await ProcessGroupService.AddHubNameAsync(NewHubName.Trim());
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Hub Added",
                        Detail = $"Hub '{NewHubName.Trim()}' added successfully.",
                        Duration = 3000
                    });
                    NewHubName = string.Empty;
                    await LoadHubNamesAsync();
                    await OnChanged.InvokeAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Add Hub",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsProcessing = false;
            }
        }

        public async Task DeleteHub(string hubName)
        {
            try
            {
                IsProcessing = true;
                var (canDelete, reason) = await ProcessGroupService.CanDeleteHubNameAsync(hubName);
                if (!canDelete)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Cannot Delete Hub",
                        Detail = reason,
                        Duration = 5000
                    });
                    return;
                }

                bool deleted = await ProcessGroupService.DeleteHubNameAsync(hubName);
                if (deleted)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Hub Deleted",
                        Detail = $"Hub '{hubName}' removed successfully.",
                        Duration = 3000
                    });
                    await LoadHubNamesAsync();
                    await OnChanged.InvokeAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Deleting Hub",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsProcessing = false;
            }
        }

        public async Task CloseModal()
        {
            NewHubName = string.Empty;
            await OnClosed.InvokeAsync();
        }
    }
}
