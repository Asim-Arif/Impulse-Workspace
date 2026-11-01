using System;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GateInReceive : ComponentBase
{
    [Inject] private IGateInDataAccess GateInDataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    [Parameter] public int EntryId { get; set; }
    [Parameter, SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

    protected GateReceiveModel? Model { get; set; }
    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; }

    protected override async Task OnInitializedAsync()
    {
        await LoadAsync();
    }

    private async Task LoadAsync()
    {
        IsLoading = true;
        try
        {
            Model = await GateInDataAccess.GetReceiveAsync(EntryId);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Error Loading", ex.Message);
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected async Task SaveAsync()
    {
        if (Model == null) return;

        var toReceive = Model.Lines.Where(x => x.ReceiveQty > 0).ToList();
        if (toReceive.Count == 0)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Notice", "Please specify a quantity to receive for at least one item.");
            return;
        }

        foreach (var line in toReceive)
        {
            if (line.ReceiveQty > line.BalanceQty)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", $"Received quantity for '{line.Detail}' cannot exceed balance quantity ({line.BalanceQty}).");
                return;
            }
        }

        IsSaving = true;
        try
        {
            await GateInDataAccess.SaveReceiveAsync(Model, Environment.UserName);
            NotificationService.Notify(NotificationSeverity.Success, "Received", "Items marked as received successfully.");
            Cancel();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Save Error", ex.Message);
        }
        finally
        {
            IsSaving = false;
        }
    }

    protected void Cancel()
    {
        var target = !string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/gate/in/list";
        Navigation.NavigateTo(target);
    }
}

