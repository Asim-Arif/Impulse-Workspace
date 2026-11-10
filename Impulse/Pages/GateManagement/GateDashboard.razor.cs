using System;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GateDashboard : ComponentBase
{
    [Inject] private IGateDashboardDataAccess DashboardDataAccess { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IUserPermissionService PermissionService { get; set; } = default!;

    protected GateDashboardModel Model { get; set; } = new();
    protected bool IsLoading { get; set; } = true;

    protected override async Task OnInitializedAsync()
    {
        await PermissionService.InitializeAsync();
        await LoadAsync();
    }

    protected bool CanAccess(string optionId)
    {
        if (PermissionService.IsAdministrator) return true;
        return PermissionService.HasOptionAccess(optionId);
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
