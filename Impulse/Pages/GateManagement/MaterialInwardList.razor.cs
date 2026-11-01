using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class MaterialInwardList : ComponentBase
{
    [Inject] private IMaterialInwardDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    protected MaterialInwardFilterModel Filter { get; set; } = new();
    protected IReadOnlyList<MaterialInwardListRowModel> Entries { get; set; } = Array.Empty<MaterialInwardListRowModel>();
    protected bool IsLoading { get; set; } = true;
    protected int? SelectedEntryId { get; set; }

    protected override async Task OnInitializedAsync()
    {
        await RefreshAsync();
    }

    protected async Task RefreshAsync()
    {
        IsLoading = true;
        try
        {
            Entries = await DataAccess.GetEntriesAsync(Filter);
        }
        catch (Exception ex)
        {
            Entries = Array.Empty<MaterialInwardListRowModel>();
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
        Navigation.NavigateTo($"/gate/material-inward/entry/{entryId}?returnUrl=/gate/material-inward");
    }

    protected async Task DeleteAsync(int entryId)
    {
        try
        {
            await DataAccess.DeleteAsync(entryId);
            NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Inward record removed successfully.");
            await RefreshAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Delete Failed", ex.Message);
        }
    }

    protected void OnContextEdit(ItemClickEventArgs e)
    {
        if (e.Data is MaterialInwardListRowModel row)
        {
            EditEntry(row.EntryID);
        }
    }

    protected async Task OnContextDelete(ItemClickEventArgs e)
    {
        if (e.Data is MaterialInwardListRowModel row)
        {
            await DeleteAsync(row.EntryID);
        }
    }
}

