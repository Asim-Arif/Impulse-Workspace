using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
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

public partial class GateInEntry : ComponentBase
{
    [Parameter] public int EntryId { get; set; }
    [SupplyParameterFromQuery] public string? returnUrl { get; set; }

    [Inject] private IGateInDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;
    [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
    [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

    protected GateInEntryModel Model { get; set; } = new();
    protected GateInDetailModel NewItem { get; set; } = new();
    protected IReadOnlyList<GateLookupOptionModel> DepartmentOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected GateInVendorLookupModel? SelectedVendor { get; set; }
    protected GateInEmployeeLookupModel? SelectedEmployee { get; set; }

    protected bool IsEditMode => EntryId > 0;
    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; } = false;

    protected override async Task OnInitializedAsync()
    {
        IsLoading = true;
        try
        {
            DepartmentOptions = await DataAccess.GetDepartmentsAsync();

            if (IsEditMode)
            {
                var existing = await DataAccess.GetGateInAsync(EntryId);
                if (existing != null)
                {
                    Model = existing;

                    if (!string.IsNullOrWhiteSpace(Model.VM_AccNo))
                    {
                        SelectedVendor = await DataAccess.GetVendorByAccNoAsync(Model.VM_AccNo);
                    }

                    if (!string.IsNullOrWhiteSpace(Model.EmpID))
                    {
                        SelectedEmployee = await DataAccess.GetEmployeeByIdAsync(Model.EmpID);
                    }
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Not Found", $"Gate In record #{EntryId} was not found.");
                    Navigation.NavigateTo(returnUrl ?? "/gate/in");
                    return;
                }
            }
            else
            {
                Model = new GateInEntryModel
                {
                    DT = DateTime.Now,
                    GatePassNo = await DataAccess.GetNextGatePassNoAsync(DateTime.Now),
                    GatePassType = GateInPassTypes.Foc,
                    GatePassType2 = GateInPassKinds.Permanent
                };
            }
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Initialization Error", ex.Message);
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
            Model.DT = dt;
            if (!IsEditMode)
            {
                Model.GatePassNo = await DataAccess.GetNextGatePassNoAsync(dt);
            }
        }
    }

    protected async Task<IEnumerable<GateInVendorLookupModel>> SearchVendorsAsync(string searchText)
    {
        return await DataAccess.SearchVendorsAsync(searchText);
    }

    protected Task OnVendorSelectedAsync(GateInVendorLookupModel? vendor)
    {
        SelectedVendor = vendor;
        Model.VM_AccNo = vendor?.AccNo;
        Model.VendorName = vendor?.AccTitle;
        return Task.CompletedTask;
    }

    protected async Task<IEnumerable<GateInEmployeeLookupModel>> SearchEmployeesAsync(string searchText)
    {
        return await DataAccess.SearchEmployeesAsync(searchText, Model.DeptID);
    }

    protected Task OnEmployeeSelectedAsync(GateInEmployeeLookupModel? emp)
    {
        SelectedEmployee = emp;
        Model.EmpID = emp?.EmpID ?? string.Empty;
        Model.EmpName = emp?.Name ?? string.Empty;
        return Task.CompletedTask;
    }

    protected void AddItem()
    {
        if (string.IsNullOrWhiteSpace(NewItem.Detail))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Missing Detail", "Please enter the item detail/description.");
            return;
        }

        if (NewItem.Qty <= 0)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Invalid Quantity", "Item quantity must be greater than zero.");
            return;
        }

        Model.Details.Add(new GateInDetailModel
        {
            Detail = NewItem.Detail.Trim(),
            Unit = !string.IsNullOrWhiteSpace(NewItem.Unit) ? NewItem.Unit.Trim() : "Pcs",
            Qty = NewItem.Qty,
            Rate = NewItem.Rate
        });

        NewItem = new GateInDetailModel();
    }

    protected void RemoveItem(GateInDetailModel item)
    {
        Model.Details.Remove(item);
    }

    protected async Task<bool> ExecuteSaveAsync()
    {
        if (string.IsNullOrWhiteSpace(Model.DeptID))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select a department.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(Model.GatePassNo))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please enter Gate Pass number.");
            return false;
        }

        if (Model.Details.Count == 0)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please add at least one material item to the pass.");
            return false;
        }

        IsSaving = true;
        try
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var userName = authState.User.Identity?.Name ?? Environment.UserName;

            var savedId = await DataAccess.SaveGateInAsync(Model, userName, Environment.MachineName);
            Model.EntryId = savedId;
            Model.EntryID = savedId;

            NotificationService.Notify(NotificationSeverity.Success, "Saved Successfully", $"Gate In Pass '{Model.GatePassNo}' has been saved.");

            if (Model.PrintSlipOnSave)
            {
                if (!Model.Approved_Finance)
                {
                    NotificationService.Notify(NotificationSeverity.Info, "Approval Required", "Pass is saved. Printing slip requires Finance Department approval.");
                }
                else
                {
                    var request = new ReportRequest
                    {
                        ReportName = ReportNames.GateManagement.MiscGateInTemp,
                        SelectionFormula = $"{{MiscEntries_GateIn.EntryID}} = {savedId}"
                    };
                    await ReportNavigation.PrintReportAsync(request);
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Save Error", ex.Message);
            return false;
        }
        finally
        {
            IsSaving = false;
        }
    }

    protected async Task SaveAsync()
    {
        var success = await ExecuteSaveAsync();
        if (success)
        {
            Navigation.NavigateTo(returnUrl ?? "/gate/in");
        }
    }

    protected async Task SaveAndNewAsync()
    {
        var success = await ExecuteSaveAsync();
        if (success)
        {
            Model = new GateInEntryModel
            {
                DT = DateTime.Now,
                GatePassNo = await DataAccess.GetNextGatePassNoAsync(DateTime.Now),
                GatePassType = GateInPassTypes.Foc,
                GatePassType2 = GateInPassKinds.Permanent
            };
            SelectedVendor = null;
            SelectedEmployee = null;
            NewItem = new GateInDetailModel();
        }
    }

    protected void Cancel()
    {
        Navigation.NavigateTo(returnUrl ?? "/gate/in");
    }
}
