using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ReportModels;
using DataAccessLibrary.Models.ViewModels.Gate;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class CourierIssuanceList : ComponentBase
{
    [Inject] private ICourierIssuanceDataAccess CourierDataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
    [Inject] private IJSRuntime JS { get; set; } = default!;

    protected CourierIssuanceFilterModel Filter { get; set; } = new();
    protected IReadOnlyList<CourierIssuanceListRowModel> Entries { get; set; } = Array.Empty<CourierIssuanceListRowModel>();
    protected IReadOnlyList<CourierVendorLookupModel> VendorOptions { get; set; } = Array.Empty<CourierVendorLookupModel>();
    protected CourierVendorLookupModel? SelectedVendor { get; set; }

    protected bool IsLoading { get; set; } = true;
    protected int? SelectedEntryId { get; set; }
    protected bool IsCustomRange => Filter.RangeKey == CourierIssuanceRangeKeys.Custom;

    protected override async Task OnInitializedAsync()
    {
        try
        {
            VendorOptions = await CourierDataAccess.GetVendorOptionsAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Vendor Load Failed", ex.Message);
        }

        await RefreshAsync();
    }

    protected async Task RefreshAsync()
    {
        IsLoading = true;
        try
        {
            if (Filter.FromDate.Date > Filter.ToDate.Date)
            {
                (Filter.FromDate, Filter.ToDate) = (Filter.ToDate, Filter.FromDate);
            }

            Entries = await CourierDataAccess.GetCourierIssuanceListAsync(Filter);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Load Failed", ex.Message);
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected async Task RangeChangedAsync(ChangeEventArgs args)
    {
        Filter.RangeKey = args.Value?.ToString() ?? CourierIssuanceRangeKeys.Last15Days;
        switch (Filter.RangeKey)
        {
            case CourierIssuanceRangeKeys.Last15Days:
                Filter.FromDate = DateTime.Today.AddDays(-15);
                Filter.ToDate = DateTime.Today;
                break;
            case CourierIssuanceRangeKeys.Last30Days:
                Filter.FromDate = DateTime.Today.AddDays(-30);
                Filter.ToDate = DateTime.Today;
                break;
            case CourierIssuanceRangeKeys.Last60Days:
                Filter.FromDate = DateTime.Today.AddDays(-60);
                Filter.ToDate = DateTime.Today;
                break;
            case CourierIssuanceRangeKeys.Last90Days:
                Filter.FromDate = DateTime.Today.AddDays(-90);
                Filter.ToDate = DateTime.Today;
                break;
            case CourierIssuanceRangeKeys.Last180Days:
                Filter.FromDate = DateTime.Today.AddDays(-180);
                Filter.ToDate = DateTime.Today;
                break;
        }
        await RefreshAsync();
    }

    protected async Task FromDateChangedAsync(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            Filter.FromDate = dt;
            await RefreshAsync();
        }
    }

    protected async Task ToDateChangedAsync(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            Filter.ToDate = dt;
            await RefreshAsync();
        }
    }

    protected Task<IEnumerable<CourierVendorLookupModel>> SearchVendorsAsync(string searchText)
    {
        if (string.IsNullOrWhiteSpace(searchText))
        {
            return Task.FromResult<IEnumerable<CourierVendorLookupModel>>(VendorOptions);
        }

        var results = VendorOptions.Where(x =>
            (!string.IsNullOrWhiteSpace(x.AccTitle) && x.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.AccNo) && x.AccNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.Address) && x.Address.Contains(searchText, StringComparison.OrdinalIgnoreCase)));

        return Task.FromResult<IEnumerable<CourierVendorLookupModel>>(results);
    }

    protected async Task OnVendorChangedAsync(CourierVendorLookupModel? vendor)
    {
        SelectedVendor = vendor;
        Filter.VendID = vendor?.AccNo;
        Filter.VendorName = vendor?.AccTitle ?? "<All Vendors>";
        await RefreshAsync();
    }

    protected async Task SearchInputAsync(ChangeEventArgs args)
    {
        Filter.SearchText = args.Value?.ToString() ?? "";
        await RefreshAsync();
    }

    protected async Task ClearSearchAsync()
    {
        Filter.SearchText = "";
        await RefreshAsync();
    }

    protected void SelectRow(int entryId) => SelectedEntryId = entryId;

    protected void NewEntry() => Navigation.NavigateTo("/gate/courier-issuance/entry?returnUrl=/gate/courier-issuance");

    protected void EditEntry(int entryId) => Navigation.NavigateTo($"/gate/courier-issuance/edit/{entryId}?returnUrl=/gate/courier-issuance");

    protected async Task DeleteEntryAsync(int entryId)
    {
        var entry = Entries.FirstOrDefault(x => x.EntryID == entryId);
        var label = entry != null && !string.IsNullOrWhiteSpace(entry.TrackingNo) 
            ? $"Courier Pass with Tracking #{entry.TrackingNo}" 
            : $"Courier Pass #{entryId}";

        var confirmed = await JS.InvokeAsync<bool>("confirm", $"Are you sure you want to delete {label}?");
        if (!confirmed) return;

        try
        {
            await CourierDataAccess.DeleteCourierIssuanceAsync(entryId);
            NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Courier issuance deleted successfully.");
            await RefreshAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Delete Failed", ex.Message);
        }
    }

    protected async Task PrintSlipAsync(int entryId)
    {
        try
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.Courier_Issuance_Slip,
                SelectionFormula = $"{{Courier_Issuance.EntryID}} = {entryId}"
            };

            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message);
        }
    }

    protected async Task PrintListAsync()
    {
        try
        {
            var from = Filter.FromDate;
            var to = Filter.ToDate;
            var selectionFormula = $"{{ItemIssuanceMisc.DT}} >= DateTime({from.Year}, {from.Month}, {from.Day}, 0, 0, 0) AND {{ItemIssuanceMisc.DT}} <= DateTime({to.Year}, {to.Month}, {to.Day}, 23, 59, 59)";
            if (!string.IsNullOrWhiteSpace(Filter.VendID))
            {
                selectionFormula += $" AND {{ItemIssuanceMisc.VendID}} = '{Filter.VendID.Replace("'", "''")}'";
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.Courier_Issuance_List,
                SelectionFormula = selectionFormula,
                FormulaValues = new Dictionary<string, object>
                {
                    { "DateRange", $"'{from:dd-MMM-yyyy} to {to:dd-MMM-yyyy}'" }
                }
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
        if (e.Data is CourierIssuanceListRowModel row)
        {
            EditEntry(row.EntryID);
        }
    }

    protected async Task OnContextPrintSlip(ItemClickEventArgs e)
    {
        if (e.Data is CourierIssuanceListRowModel row)
        {
            await PrintSlipAsync(row.EntryID);
        }
    }

    protected async Task OnContextDelete(ItemClickEventArgs e)
    {
        if (e.Data is CourierIssuanceListRowModel row)
        {
            await DeleteEntryAsync(row.EntryID);
        }
    }
}

