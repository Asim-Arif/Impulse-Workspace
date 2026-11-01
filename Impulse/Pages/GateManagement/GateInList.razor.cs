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
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GateInList : ComponentBase
{
    [Inject] private IGateInDataAccess GateInDataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;

    protected GateInFilterModel Filter { get; set; } = new();
    protected IReadOnlyList<GateInListRowModel> Entries { get; set; } = Array.Empty<GateInListRowModel>();
    protected IReadOnlyList<GateLookupOptionModel> DepartmentOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<GateLookupOptionModel> GatePassTypeOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<GateLookupOptionModel> VendorFilterOptions { get; set; } = Array.Empty<GateLookupOptionModel>();

    protected bool IsLoading { get; set; } = true;
    protected int? SelectedEntryId { get; set; }
    protected bool IsCustomRange => Filter.RangeKey == GateInRangeKeys.Custom;
    protected int TypeSelectValue => Filter.GatePassType ?? GateInPassTypes.All;

    protected override async Task OnInitializedAsync()
    {
        await LoadLookupOptionsAsync();
        await RefreshAsync();
    }

    private async Task LoadLookupOptionsAsync()
    {
        try
        {
            DepartmentOptions = await GateInDataAccess.GetDepartmentsAsync();
            GatePassTypeOptions = await GateInDataAccess.GetGatePassTypeOptionsAsync();
            VendorFilterOptions = await GateInDataAccess.GetVendorFilterOptionsAsync();
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
            if (Filter.FromDate.Date > Filter.ToDate.Date)
            {
                (Filter.FromDate, Filter.ToDate) = (Filter.ToDate, Filter.FromDate);
            }

            Entries = await GateInDataAccess.GetGateInEntriesAsync(Filter);
        }
        catch (Exception ex)
        {
            Entries = Array.Empty<GateInListRowModel>();
            NotificationService.Notify(NotificationSeverity.Error, "Error Loading List", ex.Message);
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected void SelectRow(int entryId) => SelectedEntryId = entryId;

    protected void NewEntry() => Navigation.NavigateTo("/gate/in/entry?returnUrl=/gate/in/list");

    protected void EditEntry(int entryId)
    {
        var row = Entries.FirstOrDefault(x => x.EntryId == entryId);
        if (row?.IsReceived == true)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Notice", "This gate pass is fully received and cannot be edited.");
            return;
        }
        Navigation.NavigateTo($"/gate/in/entry/{entryId}?returnUrl=/gate/in/list");
    }

    protected void ReceiveEntry(int entryId)
    {
        Navigation.NavigateTo($"/gate/in/receive/{entryId}?returnUrl=/gate/in/list");
    }

    protected async Task DepartmentChangedAsync(ChangeEventArgs args)
    {
        Filter.DepartmentId = args.Value?.ToString() ?? "";
        await RefreshAsync();
    }

    protected async Task TypeChangedAsync(ChangeEventArgs args)
    {
        var raw = args.Value?.ToString() ?? "";
        if (int.TryParse(raw, out var parsed) && parsed != GateInPassTypes.All)
        {
            Filter.GatePassType = parsed;
        }
        else
        {
            Filter.GatePassType = null;
        }
        await RefreshAsync();
    }

    protected async Task GatePassTypeNameChangedAsync(ChangeEventArgs args)
    {
        Filter.GatePassTypeName = args.Value?.ToString() ?? "<All Gatepass Types>";
        await RefreshAsync();
    }

    protected async Task RangeChangedAsync(ChangeEventArgs args)
    {
        Filter.RangeKey = args.Value?.ToString() ?? GateInRangeKeys.Last15Days;
        switch (Filter.RangeKey)
        {
            case GateInRangeKeys.Last15Days:
                Filter.FromDate = DateTime.Today.AddDays(-15);
                Filter.ToDate = DateTime.Today;
                break;
            case GateInRangeKeys.Last30Days:
                Filter.FromDate = DateTime.Today.AddDays(-30);
                Filter.ToDate = DateTime.Today;
                break;
            case GateInRangeKeys.Last60Days:
                Filter.FromDate = DateTime.Today.AddDays(-60);
                Filter.ToDate = DateTime.Today;
                break;
            case GateInRangeKeys.Last90Days:
                Filter.FromDate = DateTime.Today.AddDays(-90);
                Filter.ToDate = DateTime.Today;
                break;
            case GateInRangeKeys.Last180Days:
                Filter.FromDate = DateTime.Today.AddDays(-180);
                Filter.ToDate = DateTime.Today;
                break;
        }
        await RefreshAsync();
    }

    protected async Task VendorChangedAsync(ChangeEventArgs args)
    {
        Filter.VendorName = args.Value?.ToString() ?? "<All Vendors>";
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
                ReportName = ReportNames.GateManagement.MiscGateInTemp,
                Parameters = parameters
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
            var parameters = new Dictionary<string, object>
            {
                { "@DTFrom", Filter.FromDate.Date },
                { "@DTTo", Filter.ToDate.Date }
            };

            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.MiscGateInTempList,
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
        if (e.Data is GateInListRowModel row)
        {
            EditEntry(row.EntryId);
        }
    }

    protected void OnContextReceive(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            ReceiveEntry(row.EntryId);
        }
    }

    protected async Task OnContextPrintSlip(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            await PrintSlipAsync(row.EntryId);
        }
    }

    protected async Task OnContextApproveFinance(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            try
            {
                await GateInDataAccess.ApproveFinanceAsync(row.EntryId, Environment.UserName);
                NotificationService.Notify(NotificationSeverity.Success, "Approved", $"Gate pass {row.GatePassNo} approved by Finance.");
                await RefreshAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Approval Failed", ex.Message);
            }
        }
    }

    protected void NavigateGatePassTypes() => Navigation.NavigateTo("/gate/pass-types");
}

