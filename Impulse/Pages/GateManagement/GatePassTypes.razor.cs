using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GatePassTypes : ComponentBase
{
    [Inject] private IGatePassTypeDataAccess DataAccess { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    protected GatePassTypeModel Model { get; set; } = new();
    protected IReadOnlyList<GatePassTypeModel> Types { get; set; } = Array.Empty<GatePassTypeModel>();
    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; }

    protected override async Task OnInitializedAsync()
    {
        await LoadAsync();
    }

    protected async Task LoadAsync()
    {
        IsLoading = true;
        try
        {
            Types = await DataAccess.GetAsync();
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

    protected void NewOption()
    {
        Model = new GatePassTypeModel();
    }

    protected void Edit(GatePassTypeModel item)
    {
        Model = new GatePassTypeModel
        {
            EntryId = item.EntryId,
            GatePassType = item.GatePassType,
            Active = item.Active
        };
    }

    protected async Task SaveAsync()
    {
        if (string.IsNullOrWhiteSpace(Model.GatePassType))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Gate pass type name is required.");
            return;
        }

        IsSaving = true;
        try
        {
            await DataAccess.SaveAsync(Model);
            NotificationService.Notify(NotificationSeverity.Success, "Saved", "Gate pass type saved successfully.");
            NewOption();
            await LoadAsync();
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

    protected async Task ToggleActiveAsync(GatePassTypeModel item)
    {
        try
        {
            await DataAccess.SetActiveAsync(item.EntryId, !item.Active);
            NotificationService.Notify(NotificationSeverity.Success, "Updated", $"Type '{item.GatePassType}' status updated.");
            await LoadAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Update Error", ex.Message);
        }
    }
}

