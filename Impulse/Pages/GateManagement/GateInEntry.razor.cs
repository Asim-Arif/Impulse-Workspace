using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class GateInEntry : ComponentBase
{
    [Inject] private IGateInDataAccess GateInDataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    [Parameter] public int? EntryId { get; set; }
    [Parameter, SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

    protected GateInEntryModel Model { get; set; } = new();
    protected GateInDetailModel NewDetail { get; set; } = new();
    protected IReadOnlyList<GateLookupOptionModel> DepartmentOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<GateLookupOptionModel> EmployeeOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<GateLookupOptionModel> GatePassTypeOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<GateLookupOptionModel> VendorOptions { get; set; } = Array.Empty<GateLookupOptionModel>();

    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; }
    protected bool IsEditMode => EntryId.HasValue && EntryId.Value > 0;
    protected string DateTimeInputValue => Model.DateTime.ToString("yyyy-MM-ddTHH:mm");

    protected override async Task OnInitializedAsync()
    {
        await LoadLookupsAsync();
        if (IsEditMode)
        {
            await LoadEntryAsync(EntryId!.Value);
        }
        else
        {
            Model = new GateInEntryModel
            {
                DateTime = DateTime.Now
            };
            if (DepartmentOptions.Count > 0)
            {
                Model.DepartmentId = DepartmentOptions[0].Id;
                await LoadEmployeesForDepartmentAsync(Model.DepartmentId);
            }
            if (GatePassTypeOptions.Count > 0)
            {
                Model.GatePassTypeName = GatePassTypeOptions[0].Text;
            }
            Model.GatePassNo = await GateInDataAccess.GetNextGatePassNoAsync(Model.DateTime);
            IsLoading = false;
        }
    }

    private async Task LoadLookupsAsync()
    {
        try
        {
            DepartmentOptions = await GateInDataAccess.GetDepartmentsAsync();
            GatePassTypeOptions = await GateInDataAccess.GetGatePassTypeOptionsAsync();
            VendorOptions = await GateInDataAccess.GetVendorOptionsAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Lookup Warning", ex.Message);
        }
    }

    private async Task LoadEmployeesForDepartmentAsync(string deptId)
    {
        try
        {
            EmployeeOptions = await GateInDataAccess.GetEmployeesAsync(deptId);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Employee Warning", ex.Message);
        }
    }

    private async Task LoadEntryAsync(int entryId)
    {
        IsLoading = true;
        try
        {
            var entry = await GateInDataAccess.GetGateInAsync(entryId);
            if (entry == null)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Not Found", "Gate In entry was not found.");
                Cancel();
                return;
            }

            Model = entry;
            await LoadEmployeesForDepartmentAsync(Model.DepartmentId);
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Error Loading Entry", ex.Message);
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected async Task DepartmentChangedAsync(ChangeEventArgs args)
    {
        Model.DepartmentId = args.Value?.ToString() ?? "";
        await LoadEmployeesForDepartmentAsync(Model.DepartmentId);
    }

    protected void EmployeeChanged(ChangeEventArgs args)
    {
        Model.EmployeeId = args.Value?.ToString() ?? "";
        var emp = EmployeeOptions.FirstOrDefault(x => x.Id == Model.EmployeeId);
        if (emp != null)
        {
            Model.EmployeeText = emp.Text;
        }
    }

    protected async Task DateTimeChangedAsync(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var parsed))
        {
            Model.DateTime = parsed;
            if (!IsEditMode)
            {
                Model.GatePassNo = await GateInDataAccess.GetNextGatePassNoAsync(parsed);
            }
        }
    }

    protected void VendorChanged(ChangeEventArgs args)
    {
        if (int.TryParse(args.Value?.ToString(), out var vendId))
        {
            Model.VendorId = vendId;
            var vend = VendorOptions.FirstOrDefault(x => x.Id == vendId.ToString());
            if (vend != null)
            {
                Model.VendorText = vend.Text;
            }
        }
    }

    protected void AddDetail()
    {
        if (string.IsNullOrWhiteSpace(NewDetail.Detail))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Item detail/name is required.");
            return;
        }

        if (NewDetail.Qty <= 0)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Quantity must be greater than zero.");
            return;
        }

        Model.Details.Add(new GateInDetailModel
        {
            Detail = NewDetail.Detail.Trim(),
            Unit = NewDetail.Unit.Trim(),
            Qty = NewDetail.Qty,
            Rate = NewDetail.Rate,
            ReturnType = NewDetail.ReturnType,
            ReturnDate = NewDetail.ReturnDate
        });

        NewDetail = new GateInDetailModel();
    }

    protected void RemoveDetail(GateInDetailModel item)
    {
        Model.Details.Remove(item);
    }

    protected async Task SaveAsync()
    {
        if (string.IsNullOrWhiteSpace(Model.DepartmentId))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select a department.");
            return;
        }

        if (string.IsNullOrWhiteSpace(Model.GatePassNo))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Gate Pass Number is required.");
            return;
        }

        if (Model.Details.Count == 0)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please add at least one line item.");
            return;
        }

        IsSaving = true;
        try
        {
            await GateInDataAccess.SaveGateInAsync(Model, Environment.UserName);
            NotificationService.Notify(NotificationSeverity.Success, "Saved", $"Gate pass {Model.GatePassNo} saved successfully.");
            Cancel();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Save Error", ex.Message);
        }
        finally
        {
            IsSaving = false;
        }
    }

    protected void Cancel()
    {
        var target = !string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/gate/in/list";
        Navigation.NavigateTo(target);
    }
}

