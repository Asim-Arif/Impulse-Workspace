using System;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GateDashboard : ComponentBase
{
    [Inject] private IGateDashboardDataAccess DashboardDataAccess { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    protected GateDashboardModel Model { get; set; } = new();
    protected bool IsLoading { get; set; } = true;

    protected override async Task OnInitializedAsync()
    {
        await LoadAsync();
    }

    protected async Task LoadAsync()
    {
        IsLoading = true;
        try
        {
            Model = await DashboardDataAccess.GetDashboardAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load dashboard: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
        }
    }
}

