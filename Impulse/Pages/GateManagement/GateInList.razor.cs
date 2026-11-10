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

public partial class GateInList : ComponentBase
{
    [Inject] private IGateInDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
    [Inject] private IJSRuntime JS { get; set; } = default!;

    protected GateInFilterModel Filter { get; set; } = new();
    protected IReadOnlyList<GateInListRowModel> Entries { get; set; } = Array.Empty<GateInListRowModel>();
    protected IReadOnlyList<GateLookupOptionModel> DepartmentOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected GateInVendorLookupModel? SelectedVendor { get; set; }

    protected bool IsLoading { get; set; } = true;
    protected bool IsCustomRange => Filter.RangeKey == GateInRangeKeys.Custom;

    protected override async Task OnInitializedAsync()
    {
        try
        {
            DepartmentOptions = await DataAccess.GetDepartmentsAsync();
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

            Entries = await DataAccess.GetGateInEntriesAsync(Filter);
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
        Filter.DeptID = !string.IsNullOrWhiteSpace(raw) ? raw : string.Empty;
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
            Filter.GatePassType = GateInPassTypes.All;
        }
        await RefreshAsync();
    }

    protected async Task<IEnumerable<GateInVendorLookupModel>> SearchVendorsAsync(string searchText)
    {
        return await DataAccess.SearchVendorsAsync(searchText);
    }

    protected async Task OnVendorChangedAsync(GateInVendorLookupModel? vendor)
    {
        SelectedVendor = vendor;
        Filter.VM_AccNo = vendor?.AccNo ?? string.Empty;
        Filter.VendorName = vendor?.AccTitle ?? string.Empty;
        await RefreshAsync();
    }

    protected async Task SearchInputAsync(ChangeEventArgs args)
    {
        Filter.SearchText = args.Value?.ToString() ?? string.Empty;
        await RefreshAsync();
    }

    protected async Task ClearSearchAsync()
    {
        Filter.SearchText = string.Empty;
        await RefreshAsync();
    }

    protected void NavigateNewPass()
    {
        Navigation.NavigateTo("/gate/in/entry");
    }

    protected void OnContextEdit(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            Navigation.NavigateTo($"/gate/in/edit/{row.EntryID}");
        }
    }

    protected async Task OnContextPrintSlip(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            await PrintSlipAsync(row);
        }
    }

    protected async Task OnContextPrintDetailList(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            await PrintDetailSingleAsync(row);
        }
    }

    protected async Task OnContextApproveHOD(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            await ApproveAsync(row, "Approved_HOD", "HOD Approval");
        }
    }

    protected async Task OnContextApproveAdmin(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            await ApproveAsync(row, "Approved", "Admin Approval");
        }
    }

    protected async Task OnContextApproveFinance(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            await ApproveAsync(row, "Approved_Finance", "Finance Dept. Approval");
        }
    }

    protected async Task OnContextDelete(ItemClickEventArgs e)
    {
        if (e.Data is GateInListRowModel row)
        {
            var confirmed = await JS.InvokeAsync<bool>("confirm", $"Are you sure you want to delete Gate In Pass '{row.GatePassNo}'? This cannot be undone.");
            if (!confirmed) return;

            try
            {
                var deleted = await DataAccess.DeleteGateInAsync(row.EntryID);
                if (deleted)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Pass Deleted", $"Gate In Pass '{row.GatePassNo}' deleted successfully.");
                    await RefreshAsync();
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Warning, "Delete Failed", "The pass could not be deleted or was already removed.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Delete Error", ex.Message);
            }
        }
    }

    protected async Task ApproveAsync(GateInListRowModel row, string approvalCol, string title)
    {
        try
        {
            var success = await DataAccess.SetApprovalAsync(row.EntryID, approvalCol, true);
            if (success)
            {
                NotificationService.Notify(NotificationSeverity.Success, $"{title} Granted", $"Pass '{row.GatePassNo}' approved successfully.");
                await RefreshAsync();
            }
            else
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Approval Failed", "Could not apply approval.");
            }
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Approval Error", ex.Message);
        }
    }

    protected async Task PrintSlipAsync(GateInListRowModel row)
    {
        if (!row.Approved_Finance)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Approvals Incomplete", "Finance Department approval is required before printing the Gate Pass Slip.");
            return;
        }

        try
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.MiscGateInTemp,
                SelectionFormula = $"{{MiscEntries_GateIn.EntryID}} = {row.EntryID}"
            };

            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
        }
    }

    protected async Task PrintDetailSingleAsync(GateInListRowModel row)
    {
        try
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.MiscGateInTempList_Detail,
                SelectionFormula = $"{{MiscEntries_GateIn.EntryID}} = {row.EntryID}"
            };

            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
        }
    }

    protected async Task PrintListAsync()
    {
        try
        {
            var cond = $"{{ItemIssuanceMisc.OnlyDT}} in DateTime({Filter.FromDate:yyyy,MM,dd}) to DateTime({Filter.ToDate:yyyy,MM,dd})";

            if (!string.IsNullOrWhiteSpace(Filter.DeptID))
            {
                cond += $" and {{ItemIssuanceMisc.DeptID}} = '{Filter.DeptID}'";
            }
            if (Filter.GatePassType != GateInPassTypes.All)
            {
                cond += $" and {{ItemIssuanceMisc.GatePassType}} = {Filter.GatePassType}";
            }
            if (!string.IsNullOrWhiteSpace(Filter.VM_AccNo))
            {
                cond += $" and {{ItemIssuanceMisc.VM_AccNo}} = '{Filter.VM_AccNo}'";
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.MiscGateInTempList,
                SelectionFormula = cond,
                FormulaValues = new Dictionary<string, object>
                {
                    ["DateRange"] = $"'{Filter.FromDate:dd-MMM-yyyy} to {Filter.ToDate:dd-MMM-yyyy}'"
                }
            };

            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
        }
    }

    protected async Task PrintDetailListAllAsync()
    {
        try
        {
            var cond = $"{{ItemIssuanceMisc.OnlyDT}} in DateTime({Filter.FromDate:yyyy,MM,dd}) to DateTime({Filter.ToDate:yyyy,MM,dd})";

            if (!string.IsNullOrWhiteSpace(Filter.DeptID))
            {
                cond += $" and {{ItemIssuanceMisc.DeptID}} = '{Filter.DeptID}'";
            }
            if (Filter.GatePassType != GateInPassTypes.All)
            {
                cond += $" and {{ItemIssuanceMisc.GatePassType}} = {Filter.GatePassType}";
            }
            if (!string.IsNullOrWhiteSpace(Filter.VM_AccNo))
            {
                cond += $" and {{ItemIssuanceMisc.VM_AccNo}} = '{Filter.VM_AccNo}'";
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.MiscGateInTempList_Detail,
                SelectionFormula = cond,
                FormulaValues = new Dictionary<string, object>
                {
                    ["DateRange"] = $"'{Filter.FromDate:dd-MMM-yyyy} to {Filter.ToDate:dd-MMM-yyyy}'"
                }
            };

            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Error", ex.Message);
        }
    }
}
