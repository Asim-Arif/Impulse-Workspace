using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ReportModels;
using DataAccessLibrary.Models.ViewModels.Gate;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class VisitorList : ComponentBase
{
    [Inject] private IMiscGateInDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;

    protected MiscGateInFilterModel Filter { get; set; } = new();
    protected IReadOnlyList<MiscGateInListRowModel> Visitors { get; set; } = Array.Empty<MiscGateInListRowModel>();
    protected IReadOnlyList<GateLookupOptionModel> EmployeeFilterOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<GateLookupOptionModel> VendorFilterOptions { get; set; } = Array.Empty<GateLookupOptionModel>();

    protected bool IsLoading { get; set; } = true;
    protected int? SelectedEntryId { get; set; }
    protected bool IsCustomRange => Filter.RangeKey == MiscGateInRangeKeys.Custom;

    protected override async Task OnInitializedAsync()
    {
        await LoadLookupsAsync();
        await RefreshAsync();
    }

    private async Task LoadLookupsAsync()
    {
        try
        {
            EmployeeFilterOptions = await DataAccess.GetEmployeeFilterOptionsAsync();
            VendorFilterOptions = await DataAccess.GetVendorFilterOptionsAsync();
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
            Visitors = await DataAccess.GetVisitorsAsync(Filter);
        }
        catch (Exception ex)
        {
            Visitors = Array.Empty<MiscGateInListRowModel>();
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
        Navigation.NavigateTo($"/gate/visitors/entry/{entryId}?returnUrl=/gate/visitors");
    }

    protected async Task RangeChangedAsync(ChangeEventArgs args)
    {
        Filter.RangeKey = args.Value?.ToString() ?? MiscGateInRangeKeys.Today;
        switch (Filter.RangeKey)
        {
            case MiscGateInRangeKeys.Today:
                Filter.FromDate = DateTime.Today;
                Filter.ToDate = DateTime.Today;
                break;
            case MiscGateInRangeKeys.Last15Days:
                Filter.FromDate = DateTime.Today.AddDays(-15);
                Filter.ToDate = DateTime.Today;
                break;
            case MiscGateInRangeKeys.Last30Days:
                Filter.FromDate = DateTime.Today.AddDays(-30);
                Filter.ToDate = DateTime.Today;
                break;
            case MiscGateInRangeKeys.Last60Days:
                Filter.FromDate = DateTime.Today.AddDays(-60);
                Filter.ToDate = DateTime.Today;
                break;
            case MiscGateInRangeKeys.Last90Days:
                Filter.FromDate = DateTime.Today.AddDays(-90);
                Filter.ToDate = DateTime.Today;
                break;
            case MiscGateInRangeKeys.Last180Days:
                Filter.FromDate = DateTime.Today.AddDays(-180);
                Filter.ToDate = DateTime.Today;
                break;
        }
        await RefreshAsync();
    }

    protected async Task EmployeeChangedAsync(ChangeEventArgs args)
    {
        Filter.EmployeeName = args.Value?.ToString() ?? "";
        await RefreshAsync();
    }

    protected async Task VendorChangedAsync(ChangeEventArgs args)
    {
        Filter.VendorName = args.Value?.ToString() ?? "";
        await RefreshAsync();
    }

    protected async Task PrintSlipAsync(int entryId)
    {
        try
        {
            var parameters = new Dictionary<string, object>
            {
                { "@EntryID", entryId }
            };

            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.Visitor,
                Parameters = parameters
            };

            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message);
        }
    }

    protected void OnContextEdit(ItemClickEventArgs e)
    {
        if (e.Data is MiscGateInListRowModel row)
        {
            EditEntry(row.EntryId);
        }
    }

    protected void OnContextPass(ItemClickEventArgs e)
    {
        if (e.Data is MiscGateInListRowModel row)
        {
            Navigation.NavigateTo($"/gate/visitors/pass/{row.EntryId}");
        }
    }

    protected async Task OnContextPrintSlip(ItemClickEventArgs e)
    {
        if (e.Data is MiscGateInListRowModel row)
        {
            await PrintSlipAsync(row.EntryId);
        }
    }

    protected void OnContextGateOut(ItemClickEventArgs e)
    {
        if (e.Data is MiscGateInListRowModel row)
        {
            Navigation.NavigateTo($"/gate/visitors/gate-out?entryId={row.EntryId}");
        }
    }
}

