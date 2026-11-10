using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ReportModels;
using DataAccessLibrary.Models.ViewModels.Gate;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GatePassPermanentEntry : ComponentBase
{
    [Parameter] public int? EntryId { get; set; }
    [Parameter] [SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

    [Inject] private IGatePassPermanentDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
    [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;

    protected GatePassPermanentEntryModel Model { get; set; } = new();
    protected GatePassPermanentDetailItemModel NewItem { get; set; } = new();

    protected IReadOnlyList<GateLookupOptionModel> DepartmentOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<CourierVendorLookupModel> VendorOptions { get; set; } = Array.Empty<CourierVendorLookupModel>();
    protected IReadOnlyList<CourierEmployeeLookupModel> EmployeeOptions { get; set; } = Array.Empty<CourierEmployeeLookupModel>();

    protected CourierVendorLookupModel? SelectedVendor { get; set; }
    protected CourierEmployeeLookupModel? SelectedEmployee { get; set; }

    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; } = false;
    protected bool IsEditMode => EntryId.HasValue && EntryId.Value > 0;

    protected override async Task OnInitializedAsync()
    {
        IsLoading = true;
        try
        {
            DepartmentOptions = await DataAccess.GetDepartmentsAsync();
            VendorOptions = await DataAccess.GetVendorOptionsAsync();
            EmployeeOptions = await DataAccess.GetEmployeeOptionsAsync();

            if (IsEditMode)
            {
                var existing = await DataAccess.GetByIdAsync(EntryId!.Value);
                if (existing != null)
                {
                    Model = existing;
                    if (!string.IsNullOrWhiteSpace(Model.VM_AccNo))
                    {
                        SelectedVendor = VendorOptions.FirstOrDefault(v => string.Equals(v.AccNo, Model.VM_AccNo, StringComparison.OrdinalIgnoreCase));
                    }
                    if (!string.IsNullOrWhiteSpace(Model.EmpID))
                    {
                        SelectedEmployee = EmployeeOptions.FirstOrDefault(e => string.Equals(e.EmpID, Model.EmpID, StringComparison.OrdinalIgnoreCase));
                    }
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Warning, "Not Found", $"Gate pass #{EntryId} was not found.");
                    NavigateBack();
                    return;
                }
            }
            else
            {
                Model = new GatePassPermanentEntryModel
                {
                    DT = DateTime.Now
                };
                Model.GatePassNo = await DataAccess.GenerateGatePassNoAsync(Model.DT);
                SelectedVendor = null;
                SelectedEmployee = null;
            }
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

    protected async Task DateChangedAsync(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            var oldMonth = Model.DT.Month;
            var oldYear = Model.DT.Year;

            Model.DT = dt;

            // Regenerate pass number if month/year changed and not edit mode
            if (!IsEditMode && (dt.Month != oldMonth || dt.Year != oldYear))
            {
                try
                {
                    Model.GatePassNo = await DataAccess.GenerateGatePassNoAsync(dt);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error regenerating pass no: {ex.Message}");
                }
            }
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

    protected void OnVendorChanged(CourierVendorLookupModel? vendor)
    {
        SelectedVendor = vendor;
        if (vendor != null)
        {
            Model.VM_AccNo = vendor.AccNo;
            Model.VendNameManual = vendor.AccTitle;
        }
        else
        {
            Model.VM_AccNo = null;
        }
    }

    protected Task<IEnumerable<CourierEmployeeLookupModel>> SearchEmployeesAsync(string searchText)
    {
        if (string.IsNullOrWhiteSpace(searchText))
        {
            return Task.FromResult<IEnumerable<CourierEmployeeLookupModel>>(EmployeeOptions);
        }

        var results = EmployeeOptions.Where(x =>
            (!string.IsNullOrWhiteSpace(x.Name) && x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.EmpID) && x.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.FName) && x.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(x.Designation) && x.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase)));

        return Task.FromResult<IEnumerable<CourierEmployeeLookupModel>>(results);
    }

    protected void OnEmployeeChanged(CourierEmployeeLookupModel? employee)
    {
        SelectedEmployee = employee;
        Model.EmpID = employee?.EmpID ?? "";
    }

    protected void AddItem()
    {
        if (string.IsNullOrWhiteSpace(NewItem.Detail))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Item detail / description is required.");
            return;
        }

        if (NewItem.Qty <= 0)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Item quantity must be greater than zero.");
            return;
        }

        Model.Items.Add(new GatePassPermanentDetailItemModel
        {
            Detail = NewItem.Detail.Trim(),
            Unit = !string.IsNullOrWhiteSpace(NewItem.Unit) ? NewItem.Unit.Trim() : "Pcs",
            Qty = NewItem.Qty,
            Rate = NewItem.Rate
        });

        NewItem = new GatePassPermanentDetailItemModel
        {
            Unit = "Pcs",
            Qty = 1,
            Rate = 0
        };
    }

    protected void RemoveItem(GatePassPermanentDetailItemModel item)
    {
        Model.Items.Remove(item);
    }

    protected async Task SaveAsync()
    {
        var savedId = await ExecuteSaveAsync();
        if (savedId.HasValue)
        {
            IsSaving = true; // Keep disabled while navigating away
            if (Model.PrintSlipOnSave)
            {
                await PrintSlipAsync(savedId.Value);
            }
            NavigateBack();
        }
    }

    protected async Task SaveAndNewAsync()
    {
        var savedId = await ExecuteSaveAsync();
        if (savedId.HasValue)
        {
            if (Model.PrintSlipOnSave)
            {
                await PrintSlipAsync(savedId.Value);
            }

            Model = new GatePassPermanentEntryModel
            {
                DT = DateTime.Now
            };
            try
            {
                Model.GatePassNo = await DataAccess.GenerateGatePassNoAsync(Model.DT);
            }
            catch {}

            SelectedVendor = null;
            SelectedEmployee = null;
            NewItem = new();
            EntryId = null;
            StateHasChanged();
        }
    }

    private async Task<int?> ExecuteSaveAsync()
    {
        if (string.IsNullOrWhiteSpace(Model.GatePassNo))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Gate pass number is required.");
            return null;
        }

        if (string.IsNullOrWhiteSpace(Model.DeptID))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select a department.");
            return null;
        }

        if (Model.Items.Count == 0)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please add at least one outward item.");
            return null;
        }

        IsSaving = true;
        try
        {
            var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
            var userName = authState.User.Identity?.Name ?? Environment.UserName;
            var machineName = Environment.MachineName;

            var resultId = await DataAccess.SaveAsync(Model, userName, machineName);

            // Update in-memory model so subsequent operations recognize this record ID
            Model.EntryID = resultId;

            NotificationService.Notify(
                NotificationSeverity.Success, 
                "Saved Successfully", 
                $"Gate pass #{Model.GatePassNo} was saved successfully.");

            return resultId;
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Save Failed", ex.Message);
            return null;
        }
        finally
        {
            IsSaving = false;
        }
    }

    protected async Task PrintSlipAsync(int entryId)
    {
        try
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.GateManagement.MiscGateOut,
                SelectionFormula = $"{{MiscEntries_GateIn.EntryID}} = {entryId}"
            };

            await ReportNavigation.PrintReportAsync(request);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message);
        }
    }

    protected void Cancel()
    {
        NavigateBack();
    }

    private void NavigateBack()
    {
        if (!string.IsNullOrWhiteSpace(ReturnUrl))
        {
            Navigation.NavigateTo(ReturnUrl);
        }
        else
        {
            Navigation.NavigateTo("/gate/pass-permanent");
        }
    }
}

