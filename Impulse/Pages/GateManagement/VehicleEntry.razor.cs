using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class VehicleEntry : ComponentBase
{
    [Inject] private IVehicleManagementDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    [Parameter] public int? EntryId { get; set; }
    [Parameter, SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

    protected VehicleEntryModel Model { get; set; } = new();
    protected IReadOnlyList<VehicleCategoryModel> Categories { get; set; } = Array.Empty<VehicleCategoryModel>();

    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; }
    protected bool IsEditMode => EntryId.HasValue && EntryId.Value > 0;
    protected string EntryTimeInputValue => Model.EntryTime.ToString("yyyy-MM-ddTHH:mm");
    protected string ExitTimeInputValue => (Model.ExitTime ?? DateTime.Now).ToString("yyyy-MM-ddTHH:mm");

    protected override async Task OnInitializedAsync()
    {
        await LoadCategoriesAsync();
        if (IsEditMode)
        {
            await LoadVehicleAsync(EntryId!.Value);
        }
        else
        {
            Model = new VehicleEntryModel
            {
                EntryTime = DateTime.Now,
                Status = "Inside"
            };
            if (Categories.Count > 0)
            {
                Model.CategoryId = Categories[0].EntryID;
            }
            IsLoading = false;
        }
    }

    private async Task LoadCategoriesAsync()
    {
        try
        {
            Categories = await DataAccess.GetCategoriesAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Lookup Warning", ex.Message);
        }
    }

    private async Task LoadVehicleAsync(int id)
    {
        IsLoading = true;
        try
        {
            var entry = await DataAccess.GetVehicleAsync(id);
            if (entry == null)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Not Found", "Vehicle record was not found.");
                Cancel();
                return;
            }
            Model = entry;
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

    protected void OnEntryTimeChanged(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            Model.EntryTime = dt;
        }
    }

    protected void OnExitTimeChanged(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            Model.ExitTime = dt;
        }
    }

    protected async Task SaveAsync()
    {
        if (string.IsNullOrWhiteSpace(Model.VehicleNumber))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Vehicle Registration Number is required.");
            return;
        }

        if (Model.Status == "Closed" && !Model.ExitTime.HasValue)
        {
            Model.ExitTime = DateTime.Now;
        }

        IsSaving = true;
        try
        {
            await DataAccess.SaveVehicleAsync(Model, Environment.UserName);
            NotificationService.Notify(NotificationSeverity.Success, "Saved", $"Vehicle '{Model.VehicleNumber}' record saved successfully.");
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
        var target = !string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/gate/vehicles";
        Navigation.NavigateTo(target);
    }
}

