using BlazorContextMenu;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GatePassReturnableList : ComponentBase
{
    [Inject] protected IGatePassReturnableDataAccess DataAccess { get; set; } = default!;
    [Inject] protected NavigationManager Navigation { get; set; } = default!;
    [Inject] protected NotificationService NotificationService { get; set; } = default!;
    [Inject] protected IReportNavigationService ReportNavigation { get; set; } = default!;

    protected GatePassReturnableFilterModel Filter { get; set; } = new();
    protected List<GatePassReturnableListRowModel> Passes { get; set; } = new();
    protected List<GatePassReturnableListRowModel> FilteredPasses { get; set; } = new();

    protected IReadOnlyList<GateLookupOptionModel> GatePassTypeOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<CourierVendorLookupModel> VendorOptions { get; set; } = Array.Empty<CourierVendorLookupModel>();
    protected CourierVendorLookupModel? SelectedVendor { get; set; }

    protected GatePassReturnableListRowModel? SelectedRow { get; set; }
    protected bool IsLoading { get; set; } = true;

    // Metrics
    protected int TotalCount => Passes.Count;
    protected int OpenCount => Passes.Count(x => x.Returnable && x.TotalQtyRcvd <= 0);
    protected int PartiallyCount => Passes.Count(x => x.Returnable && x.TotalQtyRcvd > 0 && x.TotalQtyRcvd < x.TotalQty);
    protected int ClosedCount => Passes.Count(x => x.Returnable && x.TotalQtyRcvd >= x.TotalQty && x.TotalQty > 0);

    // Receiving Modal State
    protected bool IsReceiveModalVisible { get; set; }
    protected GatePassReceiveModel? ReceiveModalModel { get; set; }

    protected override async Task OnInitializedAsync()
    {
        await LoadLookupsAsync();
        ApplyPresetDates(Filter.DateRangePreset);
        await RefreshListAsync();
    }

    private async Task LoadLookupsAsync()
    {
        try
        {
            GatePassTypeOptions = await DataAccess.GetGatePassTypesAsync();
            VendorOptions = await DataAccess.GetVendorsAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Error Loading Lookups", ex.Message);
        }
    }

    protected async Task RefreshListAsync()
    {
        IsLoading = true;
        try
        {
            var results = await DataAccess.GetListAsync(Filter);
            Passes = results.ToList();
            ApplyFilter();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Error Fetching Passes", ex.Message);
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected void ApplyFilter()
    {
        if (string.IsNullOrWhiteSpace(Filter.SearchText))
        {
            FilteredPasses = Passes;
            return;
        }

        var term = Filter.SearchText.Trim();
        FilteredPasses = Passes.Where(x =>
            (!string.IsNullOrWhiteSpace(x.GatePassNo) && x.GatePassNo.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.IssdTo) && x.IssdTo.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.PersonName) && x.PersonName.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.SenderName) && x.SenderName.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.Name) && x.Name.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.DeptName_From) && x.DeptName_From.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.DeptName_To) && x.DeptName_To.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.ApprovedBy_Name) && x.ApprovedBy_Name.Contains(term, StringComparison.OrdinalIgnoreCase))
        ).ToList();
    }

    protected void ClearSearch()
    {
        Filter.SearchText = string.Empty;
        ApplyFilter();
    }

    protected async Task OnRangePresetChanged(ChangeEventArgs e)
    {
        if (int.TryParse(e.Value?.ToString(), out var preset))
        {
            Filter.DateRangePreset = preset;
            ApplyPresetDates(preset);
            if (preset != 5)
            {
                await RefreshListAsync();
            }
        }
    }

    private void ApplyPresetDates(int preset)
    {
        var today = DateTime.Today;
        Filter.DtTo = today;
        Filter.DtFrom = preset switch
        {
            0 => today.AddDays(-15),
            1 => today.AddDays(-30),
            2 => today.AddDays(-60),
            3 => today.AddDays(-90),
            4 => today.AddDays(-180),
            _ => Filter.DtFrom
        };
    }

    protected async Task OnCustomDateFromChanged(string? value)
    {
        if (DateTime.TryParse(value, out var dt))
        {
            Filter.DtFrom = dt;
            await RefreshListAsync();
        }
    }

    protected async Task OnCustomDateToChanged(string? value)
    {
        if (DateTime.TryParse(value, out var dt))
        {
            Filter.DtTo = dt;
            await RefreshListAsync();
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

    protected async Task OnVendorChanged(CourierVendorLookupModel? vendor)
    {
        SelectedVendor = vendor;
        Filter.VM_AccNo = vendor?.AccNo ?? string.Empty;
        await RefreshListAsync();
    }

    protected void NavigateToNewEntry()
    {
        Navigation.NavigateTo("/gate/pass-returnable/entry");
    }

    protected void NavigateToEdit(int entryId)
    {
        Navigation.NavigateTo($"/gate/pass-returnable/edit/{entryId}");
    }

    protected bool CanReceive(GatePassReturnableListRowModel? row)
    {
        if (row == null) return false;
        return row.Returnable && (row.TotalQtyRcvd < row.TotalQty || row.TotalQty <= 0);
    }

    // Context Menu Handlers
    protected void OnContextEdit(ItemClickEventArgs e)
    {
        if (e.Data is GatePassReturnableListRowModel row)
        {
            NavigateToEdit(row.EntryID);
        }
    }

    protected async Task OnContextReceive(ItemClickEventArgs e)
    {
        if (e.Data is GatePassReturnableListRowModel row)
        {
            if (!row.Returnable)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Non-Returnable", "This gate pass is marked as non-returnable.");
                return;
            }

            if (row.TotalQtyRcvd >= row.TotalQty && row.TotalQty > 0)
            {
                NotificationService.Notify(NotificationSeverity.Info, "Fully Received", "All items for this gate pass have already been received.");
                return;
            }

            await OpenReceiveDialogAsync(row.EntryID);
        }
    }

    private async Task OpenReceiveDialogAsync(int entryId)
    {
        try
        {
            ReceiveModalModel = await DataAccess.GetReceiveModelByIdAsync(entryId);
            if (ReceiveModalModel != null)
            {
                IsReceiveModalVisible = true;
            }
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Error Opening Receiving", ex.Message);
        }
    }

    protected void OnReceiveModalVisibilityChanged(bool visible)
    {
        IsReceiveModalVisible = visible;
    }

    protected async Task OnReceivedSuccessfullyAsync()
    {
        await RefreshListAsync();
    }

    protected async Task OnContextPrintSlipIssuance(ItemClickEventArgs e)
    {
        if (e.Data is GatePassReturnableListRowModel row)
        {
            await PrintSlipIssuanceAsync(row.EntryID);
        }
    }

    protected async Task OnContextPrintSlipReceiving(ItemClickEventArgs e)
    {
        if (e.Data is GatePassReturnableListRowModel row)
        {
            if (row.TotalQtyRcvd <= 0)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "No Receiving", "No items have been received for this gate pass yet.");
                return;
            }

            await PrintSlipReceivingAsync(row.EntryID);
        }
    }

    private async Task PrintSlipIssuanceAsync(int entryId)
    {
        try
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.GatePass_Return,
                SelectionFormula = $"{{MiscEntries_GateIn.EntryID}} = {entryId}"
            };
            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message);
        }
    }

    private async Task PrintSlipReceivingAsync(int entryId)
    {
        try
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.GatePass_Return_Receiving,
                SelectionFormula = $"{{MiscEntries_GateIn.EntryID}} = {entryId}"
            };
            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message);
        }
    }

    // Top Options Menu Actions
    protected async Task OnContextPrintPassList(ItemClickEventArgs e)
    {
        await PrintListInternalAsync(ReportNames.GateManagement.GatePass_Return_List);
    }

    protected async Task OnContextPrintDetailList(ItemClickEventArgs e)
    {
        await PrintListInternalAsync(ReportNames.GateManagement.GatePass_Return_List_Detail);
    }

    protected async Task OnContextRefreshList(ItemClickEventArgs e)
    {
        await RefreshListAsync();
    }

    protected void OnContextNewEntry(ItemClickEventArgs e)
    {
        NavigateToNewEntry();
    }

    private async Task PrintListInternalAsync(string reportName)
    {
        try
        {
            var dFrom = Filter.DtFrom.Date;
            var dTo = Filter.DtTo.Date;

            var formula = $"{{ItemIssuanceMisc.OnlyDT}} >= DateTime({dFrom.Year}, {dFrom.Month}, {dFrom.Day}, 0, 0, 0) AND {{ItemIssuanceMisc.OnlyDT}} <= DateTime({dTo.Year}, {dTo.Month}, {dTo.Day}, 23, 59, 59)";

            if (!string.IsNullOrWhiteSpace(Filter.GatePassType))
            {
                formula += $" AND {{ItemIssuanceMisc.GatePassType}} = '{Filter.GatePassType.Trim()}'";
            }

            if (string.Equals(Filter.ReturnableFilter, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                formula += " AND {ItemIssuanceMisc.Returnable} = true";
            }
            else if (string.Equals(Filter.ReturnableFilter, "No", StringComparison.OrdinalIgnoreCase))
            {
                formula += " AND {ItemIssuanceMisc.Returnable} = false";
            }

            if (!string.IsNullOrWhiteSpace(Filter.VM_AccNo))
            {
                formula += $" AND {{ItemIssuanceMisc.VM_AccNo}} = '{Filter.VM_AccNo.Trim()}'";
            }

            var request = new ReportRequest
            {
                ReportName = reportName,
                SelectionFormula = formula,
                FormulaValues = new Dictionary<string, object>
                {
                    ["DateRange"] = $"{dFrom:dd-MMM-yyyy} to {dTo:dd-MMM-yyyy}"
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

