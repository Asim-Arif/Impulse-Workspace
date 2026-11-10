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

public partial class GatePassPermanentList : ComponentBase
{
    [Inject] private IGatePassPermanentDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
    [Inject] private IJSRuntime JS { get; set; } = default!;

    protected GatePassPermanentFilterModel Filter { get; set; } = new();
    protected IReadOnlyList<GatePassPermanentListRowModel> Entries { get; set; } = Array.Empty<GatePassPermanentListRowModel>();
    protected IReadOnlyList<GateLookupOptionModel> DepartmentOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<CourierVendorLookupModel> VendorOptions { get; set; } = Array.Empty<CourierVendorLookupModel>();
    protected CourierVendorLookupModel? SelectedVendor { get; set; }

    protected bool IsLoading { get; set; } = true;
    protected bool IsCustomRange => Filter.RangeKey == GatePassPermanentRangeKeys.Custom;

    protected override async Task OnInitializedAsync()
    {
        try
        {
            DepartmentOptions = await DataAccess.GetDepartmentsAsync();
            VendorOptions = await DataAccess.GetVendorOptionsAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Lookup Load Failed", ex.Message);
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

            Entries = await DataAccess.GetListAsync(Filter);
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
        Filter.RangeKey = args.Value?.ToString() ?? GatePassPermanentRangeKeys.Last15Days;
        switch (Filter.RangeKey)
        {
            case GatePassPermanentRangeKeys.Last15Days:
                Filter.FromDate = DateTime.Today.AddDays(-15);
                Filter.ToDate = DateTime.Today;
                break;
            case GatePassPermanentRangeKeys.Last30Days:
                Filter.FromDate = DateTime.Today.AddDays(-30);
                Filter.ToDate = DateTime.Today;
                break;
            case GatePassPermanentRangeKeys.Last60Days:
                Filter.FromDate = DateTime.Today.AddDays(-60);
                Filter.ToDate = DateTime.Today;
                break;
            case GatePassPermanentRangeKeys.Last90Days:
                Filter.FromDate = DateTime.Today.AddDays(-90);
                Filter.ToDate = DateTime.Today;
                break;
            case GatePassPermanentRangeKeys.Last180Days:
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

    protected async Task DepartmentChangedAsync(ChangeEventArgs args)
    {
        var raw = args.Value?.ToString() ?? "";
        Filter.DeptID = !string.IsNullOrWhiteSpace(raw) ? raw : null;
        await RefreshAsync();
    }

    protected async Task TypeChangedAsync(ChangeEventArgs args)
    {
        var raw = args.Value?.ToString() ?? "";
        if (int.TryParse(raw, out var parsed))
        {
            Filter.GatePassType = parsed;
        }
        else
        {
            Filter.GatePassType = null;
        }
        await RefreshAsync();
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
        Filter.VM_AccNo = vendor?.AccNo;
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

    protected void NavigateNewPass()
    {
        Navigation.NavigateTo("/gate/pass-permanent/entry");
    }

    protected void OnContextEdit(ItemClickEventArgs e)
    {
        if (e.Data is GatePassPermanentListRowModel row)
        {
            Navigation.NavigateTo($"/gate/pass-permanent/edit/{row.EntryID}");
        }
    }

    protected async Task OnContextPrintSlip(ItemClickEventArgs e)
    {
        if (e.Data is GatePassPermanentListRowModel row)
        {
            await PrintSlipAsync(row);
        }
    }

    protected async Task OnContextPrintDetailList(ItemClickEventArgs e)
    {
        if (e.Data is GatePassPermanentListRowModel row)
        {
            await PrintDetailListAsync(row);
        }
    }

    protected async Task OnContextApproveHOD(ItemClickEventArgs e)
    {
        if (e.Data is GatePassPermanentListRowModel row)
        {
            await ApproveAsync(row, "Approved_HOD", "HOD Approval");
        }
    }

    protected async Task OnContextApproveAdmin(ItemClickEventArgs e)
    {
        if (e.Data is GatePassPermanentListRowModel row)
        {
            await ApproveAsync(row, "Approved", "Admin Approval");
        }
    }

    protected async Task OnContextApproveFinance(ItemClickEventArgs e)
    {
        if (e.Data is GatePassPermanentListRowModel row)
        {
            await ApproveAsync(row, "Approved_Finance", "Finance Dept. Approval");
        }
    }

    private async Task ApproveAsync(GatePassPermanentListRowModel row, string column, string label)
    {
        try
        {
            await DataAccess.SetApprovalAsync(row.EntryID, column, true);
            NotificationService.Notify(NotificationSeverity.Success, "Approved", $"{label} recorded for Pass #{row.GatePassNo}.");
            await RefreshAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Approval Failed", ex.Message);
        }
    }

    protected async Task OnContextDelete(ItemClickEventArgs e)
    {
        if (e.Data is GatePassPermanentListRowModel row)
        {
            var confirmed = await JS.InvokeAsync<bool>("confirm", $"Are you sure you want to delete Permanent Gate Pass #{row.GatePassNo}?");
            if (!confirmed) return;

            try
            {
                await DataAccess.DeleteAsync(row.EntryID);
                NotificationService.Notify(NotificationSeverity.Success, "Deleted", $"Gate pass #{row.GatePassNo} was deleted.");
                await RefreshAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Delete Failed", ex.Message);
            }
        }
    }

    protected async Task PrintSlipAsync(GatePassPermanentListRowModel row)
    {
        if (!row.Approved_Finance)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Approvals Incomplete", "Print Slip requires Finance Department Approval.");
            return;
        }

        try
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.MiscGateOut,
                SelectionFormula = $"{{MiscEntries_GateIn.EntryID}} = {row.EntryID}"
            };

            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message);
        }
    }

    protected async Task PrintDetailListAsync(GatePassPermanentListRowModel row)
    {
        try
        {
            var from = Filter.FromDate;
            var to = Filter.ToDate;
            var cond = $"{{ItemIssuanceMisc.DT}} >= DateTime({from.Year}, {from.Month}, {from.Day}, 0, 0, 0) AND {{ItemIssuanceMisc.DT}} <= DateTime({to.Year}, {to.Month}, {to.Day}, 23, 59, 59)";
            if (!string.IsNullOrWhiteSpace(Filter.DeptID))
            {
                cond += $" AND {{ItemIssuanceMisc.DeptID}} = '{Filter.DeptID.Replace("'", "''")}'";
            }
            if (Filter.GatePassType.HasValue)
            {
                cond += $" AND {{ItemIssuanceMisc.GatePassType}} = {Filter.GatePassType.Value}";
            }
            if (!string.IsNullOrWhiteSpace(Filter.VM_AccNo))
            {
                cond += $" AND {{ItemIssuanceMisc.VM_AccNo}} = '{Filter.VM_AccNo.Replace("'", "''")}'";
            }

            var request = new ReportRequest
            {
                ReportName = "MiscGateOutList_Detail.rpt",
                SelectionFormula = cond,
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

    protected async Task PrintListAsync()
    {
        try
        {
            var from = Filter.FromDate;
            var to = Filter.ToDate;
            var cond = $"{{ItemIssuanceMisc.DT}} >= DateTime({from.Year}, {from.Month}, {from.Day}, 0, 0, 0) AND {{ItemIssuanceMisc.DT}} <= DateTime({to.Year}, {to.Month}, {to.Day}, 23, 59, 59)";
            if (!string.IsNullOrWhiteSpace(Filter.DeptID))
            {
                cond += $" AND {{ItemIssuanceMisc.DeptID}} = '{Filter.DeptID.Replace("'", "''")}'";
            }
            if (Filter.GatePassType.HasValue)
            {
                cond += $" AND {{ItemIssuanceMisc.GatePassType}} = {Filter.GatePassType.Value}";
            }
            if (!string.IsNullOrWhiteSpace(Filter.VM_AccNo))
            {
                cond += $" AND {{ItemIssuanceMisc.VM_AccNo}} = '{Filter.VM_AccNo.Replace("'", "''")}'";
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.MiscGateOutList,
                SelectionFormula = cond,
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
}

