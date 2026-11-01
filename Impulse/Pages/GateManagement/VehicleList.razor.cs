using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class VehicleList : ComponentBase
{
    [Inject] private IVehicleManagementDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    protected VehicleFilterModel Filter { get; set; } = new();
    protected IReadOnlyList<VehicleEntryListRowModel> Vehicles { get; set; } = Array.Empty<VehicleEntryListRowModel>();
    protected IReadOnlyList<VehicleCategoryModel> Categories { get; set; } = Array.Empty<VehicleCategoryModel>();
    protected bool IsLoading { get; set; } = true;
    protected int? SelectedEntryId { get; set; }

    protected override async Task OnInitializedAsync()
    {
        await LoadCategoriesAsync();
        await RefreshAsync();
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

    protected async Task RefreshAsync()
    {
        IsLoading = true;
        try
        {
            Vehicles = await DataAccess.GetVehiclesAsync(Filter);
        }
        catch (Exception ex)
        {
            Vehicles = Array.Empty<VehicleEntryListRowModel>();
            NotificationService.Notify(NotificationSeverity.Error, "Error Loading", ex.Message);
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected void SelectRow(int entryId) => SelectedEntryId = entryId;

    protected void EditEntry(int entryId)
    {
        Navigation.NavigateTo($"/gate/vehicles/entry/{entryId}?returnUrl=/gate/vehicles");
    }

    protected async Task CategoryChangedAsync(ChangeEventArgs args)
    {
        if (int.TryParse(args.Value?.ToString(), out var catId))
        {
            Filter.CategoryId = catId;
            await RefreshAsync();
        }
    }

    protected async Task DeleteAsync(int entryId)
    {
        try
        {
            await DataAccess.DeleteVehicleAsync(entryId);
            NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Vehicle record deleted.");
            await RefreshAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Delete Failed", ex.Message);
        }
    }

    protected void OnContextEdit(ItemClickEventArgs e)
    {
        if (e.Data is VehicleEntryListRowModel row)
        {
            EditEntry(row.EntryID);
        }
    }

    protected async Task OnContextDelete(ItemClickEventArgs e)
    {
        if (e.Data is VehicleEntryListRowModel row)
        {
            await DeleteAsync(row.EntryID);
        }
    }
}

