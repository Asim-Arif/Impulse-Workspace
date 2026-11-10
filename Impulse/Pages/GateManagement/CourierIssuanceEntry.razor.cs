using System;
using System.Collections.Generic;
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

public partial class CourierIssuanceEntry : ComponentBase
{
    [Parameter] public int? EntryId { get; set; }
    [Parameter] [SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

    [Inject] private ICourierIssuanceDataAccess CourierDataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
    [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;

    protected CourierIssuanceEntryModel Model { get; set; } = new();
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
            VendorOptions = await CourierDataAccess.GetVendorOptionsAsync();
            EmployeeOptions = await CourierDataAccess.GetEmployeeOptionsAsync();

            if (IsEditMode)
            {
                var existing = await CourierDataAccess.GetCourierIssuanceAsync(EntryId!.Value);
                if (existing != null)
                {
                    Model = existing;
                    if (!string.IsNullOrWhiteSpace(Model.VendID))
                    {
                        SelectedVendor = VendorOptions.FirstOrDefault(v => string.Equals(v.AccNo, Model.VendID, StringComparison.OrdinalIgnoreCase));
                    }
                    if (!string.IsNullOrWhiteSpace(Model.IssuedBy))
                    {
                        SelectedEmployee = EmployeeOptions.FirstOrDefault(e => string.Equals(e.EmpID, Model.IssuedBy, StringComparison.OrdinalIgnoreCase));
                    }
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Warning, "Not Found", $"Courier pass #{EntryId} was not found.");
                    NavigateBack();
                    return;
                }
            }
            else
            {
                Model = new CourierIssuanceEntryModel
                {
                    DT = DateTime.Now
                };
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

    protected void DateChanged(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            Model.DT = dt;
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
            Model.VendID = vendor.AccNo;
            Model.VendorName = vendor.AccTitle;
            Model.VendorAddress = vendor.Address;
        }
        else
        {
            Model.VendID = null;
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
        Model.IssuedBy = employee?.EmpID ?? "";
    }

    protected async Task SaveAsync()
    {
        var savedId = await ExecuteSaveAsync();
        if (savedId.HasValue)
        {
            IsSaving = true;
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

            Model = new CourierIssuanceEntryModel
            {
                DT = DateTime.Now
            };
            SelectedVendor = null;
            SelectedEmployee = null;
            EntryId = null;
            StateHasChanged();
        }
    }

    private async Task<int?> ExecuteSaveAsync()
    {
        if (string.IsNullOrWhiteSpace(Model.VendorName))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Vendor / Recipient name is required.");
            return null;
        }

        IsSaving = true;
        try
        {
            var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
            var userName = authState.User.Identity?.Name ?? Environment.UserName;
            var machineName = Environment.MachineName;

            var resultId = await CourierDataAccess.SaveCourierIssuanceAsync(Model, userName, machineName);

            Model.EntryID = resultId;

            NotificationService.Notify(
                NotificationSeverity.Success, 
                "Saved", 
                IsEditMode ? $"Courier pass #{resultId} updated successfully." : $"Courier pass #{resultId} created successfully."
            );

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

    protected void Cancel() => NavigateBack();

    private void NavigateBack()
    {
        if (!string.IsNullOrWhiteSpace(ReturnUrl))
        {
            Navigation.NavigateTo(ReturnUrl);
        }
        else
        {
            Navigation.NavigateTo("/gate/courier-issuance");
        }
    }
}

