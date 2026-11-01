using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class VisitorEntry : ComponentBase
{
    [Inject] private IMiscGateInDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    [Parameter] public int? EntryId { get; set; }
    [Parameter, SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

    protected MiscGateInEntryModel Model { get; set; } = new();
    protected IReadOnlyList<GateLookupOptionModel> EmployeeOptions { get; set; } = Array.Empty<GateLookupOptionModel>();
    protected IReadOnlyList<GateLookupOptionModel> VendorOptions { get; set; } = Array.Empty<GateLookupOptionModel>();

    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; }
    protected bool IsEditMode => EntryId.HasValue && EntryId.Value > 0;
    protected string DateTimeInputValue => Model.DateTime.ToString("yyyy-MM-ddTHH:mm");
    protected string TimeOutInputValue => (Model.TimeOut ?? DateTime.Now).ToString("yyyy-MM-ddTHH:mm");

    protected override async Task OnInitializedAsync()
    {
        await LoadLookupsAsync();
        if (IsEditMode)
        {
            await LoadEntryAsync(EntryId!.Value);
        }
        else
        {
            Model = new MiscGateInEntryModel
            {
                DateTime = DateTime.Now
            };
            IsLoading = false;
        }
    }

    private async Task LoadLookupsAsync()
    {
        try
        {
            EmployeeOptions = await DataAccess.GetEmployeeOptionsAsync();
            VendorOptions = await DataAccess.GetVendorOptionsAsync();
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Lookup Warning", ex.Message);
        }
    }

    private async Task LoadEntryAsync(int id)
    {
        IsLoading = true;
        try
        {
            var entry = await DataAccess.GetVisitorAsync(id);
            if (entry == null)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Not Found", "Visitor entry not found.");
                Cancel();
                return;
            }
            Model = entry;
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Error Loading", ex.Message);
        }
        finally
        {
            IsLoading = false;
        }
    }

    protected void OnDateTimeChanged(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            Model.DateTime = dt;
        }
    }

    protected void OnTimeOutChanged(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            Model.TimeOut = dt;
        }
    }

    protected async Task OnEmployeeChangedAsync(ChangeEventArgs args)
    {
        Model.EmployeeId = args.Value?.ToString() ?? "";
        var emp = EmployeeOptions.FirstOrDefault(x => x.Id == Model.EmployeeId);
        if (emp != null)
        {
            Model.EmployeeText = emp.Text;
            var contact = await DataAccess.GetConcernedPersonContactAsync(Model.EmployeeId);
            if (contact != null)
            {
                if (!string.IsNullOrWhiteSpace(contact.EmailAddress)) Model.ConcernedPersonEmail = contact.EmailAddress;
                if (!string.IsNullOrWhiteSpace(contact.WhatsAppNo)) Model.ConcernedPersonWhatsApp = contact.WhatsAppNo;
            }
        }
    }

    protected void OnVendorChanged(ChangeEventArgs args)
    {
        if (int.TryParse(args.Value?.ToString(), out var vId))
        {
            Model.VendorId = vId;
            var v = VendorOptions.FirstOrDefault(x => x.Id == vId.ToString());
            if (v != null) Model.VendorText = v.Text;
        }
    }

    protected async Task SaveAsync()
    {
        if (string.IsNullOrWhiteSpace(Model.VisitorName))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Visitor name is required.");
            return;
        }

        IsSaving = true;
        try
        {
            var savedId = await DataAccess.SaveVisitorAsync(Model, Environment.UserName);
            NotificationService.Notify(NotificationSeverity.Success, "Saved", $"Visitor '{Model.VisitorName}' registered successfully.");
            
            // If new entry, offer pass viewing or redirect
            if (!IsEditMode)
            {
                Navigation.NavigateTo($"/gate/visitors/pass/{savedId}");
            }
            else
            {
                Cancel();
            }
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
        var target = !string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/gate/visitors";
        Navigation.NavigateTo(target);
    }
}

