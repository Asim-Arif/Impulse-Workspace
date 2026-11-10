using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GatePassReturnableEntry : ComponentBase
{
    [Inject] protected IGatePassReturnableDataAccess DataAccess { get; set; } = default!;
    [Inject] protected NavigationManager Navigation { get; set; } = default!;
    [Inject] protected NotificationService NotificationService { get; set; } = default!;
    [Inject] protected IReportNavigationService ReportNavigation { get; set; } = default!;
    [Inject] protected AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;

    [Parameter] public int? EntryId { get; set; }
    [Parameter, SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

    protected GatePassReturnableEntryModel Model { get; set; } = new();
    protected GatePassReturnableDetailItemModel NewItem { get; set; } = new();

    protected IReadOnlyList<GateLookupOptionModel> DepartmentOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<GateLookupOptionModel> GatePassTypeOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<CourierEmployeeLookupModel> EmployeeOptions { get; set; } = Array.Empty<CourierEmployeeLookupModel>();
    protected IReadOnlyList<CourierVendorLookupModel> VendorOptions { get; set; } = Array.Empty<CourierVendorLookupModel>();

    protected CourierVendorLookupModel? SelectedVendor { get; set; }
    protected CourierEmployeeLookupModel? SelectedDeptEmployee { get; set; }
    protected CourierEmployeeLookupModel? SelectedApprovedByEmployee { get; set; }

    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; }
    protected bool IsEditMode => EntryId.HasValue && EntryId.Value > 0;

    protected override async Task OnInitializedAsync()
    {
        await LoadLookupsAsync();

        if (IsEditMode)
        {
            await LoadEntryAsync(EntryId!.Value);
        }
        else
        {
            Model = new GatePassReturnableEntryModel
            {
                DT = DateTime.Now,
                Returnable = true,
                ReturnDT = DateTime.Today.AddDays(7)
            };

            try
            {
                Model.GatePassNo = await DataAccess.GenerateGatePassNoAsync(Model.DT);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating pass no: {ex.Message}");
            }

            if (GatePassTypeOptions.Count > 0)
            {
                Model.GatePassType = GatePassTypeOptions[0].Id;
            }
            if (DepartmentOptions.Count > 0)
            {
                Model.DeptID_From = DepartmentOptions[0].Id;
            }

            IsLoading = false;
        }
    }

    private async Task LoadLookupsAsync()
    {
        try
        {
            DepartmentOptions = await DataAccess.GetDepartmentsAsync();
            GatePassTypeOptions = await DataAccess.GetGatePassTypesAsync();
            EmployeeOptions = await DataAccess.GetEmployeesAsync();
            VendorOptions = await DataAccess.GetVendorsAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Error Loading Lookups", ex.Message);
        }
    }

    private async Task LoadEntryAsync(int entryId)
    {
        IsLoading = true;
        try
        {
            var existing = await DataAccess.GetByIdAsync(entryId);
            if (existing == null)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Not Found", $"Returnable gate pass #{entryId} was not found.");
                NavigateBack();
                return;
            }

            Model = existing;

            if (!string.IsNullOrWhiteSpace(Model.VM_AccNo))
            {
                SelectedVendor = VendorOptions.FirstOrDefault(x => string.Equals(x.AccNo, Model.VM_AccNo, StringComparison.OrdinalIgnoreCase));
            }

            if (!string.IsNullOrWhiteSpace(Model.DeptID_To))
            {
                SelectedDeptEmployee = EmployeeOptions.FirstOrDefault(x => string.Equals(x.EmpID, Model.DeptID_To, StringComparison.OrdinalIgnoreCase));
            }

            if (!string.IsNullOrWhiteSpace(Model.ApprovedBy_Name))
            {
                SelectedApprovedByEmployee = EmployeeOptions.FirstOrDefault(x => string.Equals(x.EmpID, Model.ApprovedBy_Name, StringComparison.OrdinalIgnoreCase));
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

    protected async Task DateChangedAsync(ChangeEventArgs e)
    {
        if (DateTime.TryParse(e.Value?.ToString(), out var dt))
        {
            var oldDay = Model.DT.Day;
            var oldMonth = Model.DT.Month;
            var oldYear = Model.DT.Year;

            Model.DT = dt;

            // Regenerate pass number if day/month/year changed and not edit mode
            if (!IsEditMode && (dt.Day != oldDay || dt.Month != oldMonth || dt.Year != oldYear))
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

    protected void OnReturnDateChanged(string? value)
    {
        if (DateTime.TryParse(value, out var dt))
        {
            Model.ReturnDT = dt;
        }
        else
        {
            Model.ReturnDT = null;
        }
    }

    protected void OnDeptFromChanged(ChangeEventArgs e)
    {
        Model.DeptID_From = e.Value?.ToString() ?? "";
    }

    protected Task<IEnumerable<CourierEmployeeLookupModel>> SearchDeptEmployeesAsync(string searchText)
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

    protected void OnDeptEmployeeChanged(CourierEmployeeLookupModel? employee)
    {
        SelectedDeptEmployee = employee;
        Model.DeptID_To = employee?.EmpID;
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
        Model.VM_AccNo = vendor?.AccNo;
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

    protected void OnApprovedByEmployeeChanged(CourierEmployeeLookupModel? employee)
    {
        SelectedApprovedByEmployee = employee;
        Model.ApprovedBy_Name = employee?.EmpID;
        Model.ApprovedBy_Designation = employee?.Designation;
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

        Model.Items.Add(new GatePassReturnableDetailItemModel
        {
            Detail = NewItem.Detail.Trim(),
            Unit = !string.IsNullOrWhiteSpace(NewItem.Unit) ? NewItem.Unit.Trim() : "Pcs",
            Qty = NewItem.Qty,
            Packing = NewItem.Packing?.Trim()
        });

        NewItem = new GatePassReturnableDetailItemModel
        {
            Unit = "Pcs",
            Qty = 1
        };
    }

    protected void RemoveItem(GatePassReturnableDetailItemModel item)
    {
        Model.Items.Remove(item);
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

            Model = new GatePassReturnableEntryModel
            {
                DT = DateTime.Now,
                Returnable = true,
                ReturnDT = DateTime.Today.AddDays(7),
                GatePassType = GatePassTypeOptions.FirstOrDefault()?.Id ?? "",
                DeptID_From = DepartmentOptions.FirstOrDefault()?.Id ?? ""
            };

            try
            {
                Model.GatePassNo = await DataAccess.GenerateGatePassNoAsync(Model.DT);
            }
            catch {}

            SelectedVendor = null;
            SelectedDeptEmployee = null;
            SelectedApprovedByEmployee = null;
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

        if (string.IsNullOrWhiteSpace(Model.DeptID_From))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select issuance from department.");
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

            // Update model in memory to track saved id
            Model.EntryID = resultId;

            NotificationService.Notify(
                NotificationSeverity.Success, 
                "Saved Successfully", 
                $"Returnable gate pass #{Model.GatePassNo} saved successfully.");

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
            Navigation.NavigateTo("/gate/pass-returnable");
        }
    }
}

