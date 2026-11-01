using System;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class VisitorPass : ComponentBase
{
    [Inject] private IMiscGateInDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IJSRuntime JS { get; set; } = default!;

    [Parameter] public int EntryId { get; set; }

    protected VisitorPassCardModel? Pass { get; set; }
    protected bool IsLoading { get; set; } = true;

    protected override async Task OnInitializedAsync()
    {
        await LoadPassAsync();
    }

    private async Task LoadPassAsync()
    {
        IsLoading = true;
        try
        {
            Pass = await DataAccess.EnsurePassAsync(EntryId);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Pass Error", ex.Message);
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected async Task PrintBrowser()
    {
        await JS.InvokeVoidAsync("window.print");
    }

    protected void Back() => Navigation.NavigateTo("/gate/visitors");
}

