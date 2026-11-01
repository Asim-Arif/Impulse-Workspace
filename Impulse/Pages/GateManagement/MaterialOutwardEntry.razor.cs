using System;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class MaterialOutwardEntry : ComponentBase
{
    [Inject] private IMaterialOutwardDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    [Parameter] public int? EntryId { get; set; }
    [Parameter, SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

    protected MaterialOutwardEntryModel Model { get; set; } = new();
    protected MaterialOutwardLineModel NewLine { get; set; } = new();
    protected bool IsLoading { get; set; } = true;
    protected bool IsSaving { get; set; }
    protected bool IsEditMode => EntryId.HasValue && EntryId.Value > 0;
    protected string DateTimeInputValue => Model.EntryTime.ToString("yyyy-MM-ddTHH:mm");

    protected override async Task OnInitializedAsync()
    {
        if (IsEditMode)
        {
            await LoadEntryAsync(EntryId!.Value);
        }
        else
        {
            Model = new MaterialOutwardEntryModel
            {
                EntryTime = DateTime.Now
            };
            IsLoading = false;
        }
    }

    private async Task LoadEntryAsync(int id)
    {
        IsLoading = true;
        try
        {
            var entry = await DataAccess.GetEntryAsync(id);
            if (entry == null)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Not Found", "Material outward entry not found.");
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
            Model.EntryTime = dt;
        }
    }

    protected void AddLine()
    {
        if (string.IsNullOrWhiteSpace(NewLine.MaterialName))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Material name/description is required.");
            return;
        }

        if (NewLine.Quantity <= 0)
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Quantity must be greater than zero.");
            return;
        }

        Model.Lines.Add(new MaterialOutwardLineModel
        {
            MaterialName = NewLine.MaterialName.Trim(),
            Quantity = NewLine.Quantity,
            Uom = NewLine.Uom?.Trim() ?? ""
        });

        NewLine = new MaterialOutwardLineModel();
    }

    protected void RemoveLine(MaterialOutwardLineModel line)
    {
        Model.Lines.Remove(line);
    }

    protected async Task SaveAsync()
    {
        if (string.IsNullOrWhiteSpace(Model.CustomerName))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Customer name is required.");
            return;
        }

        IsSaving = true;
        try
        {
            await DataAccess.SaveAsync(Model, Environment.UserName);
            NotificationService.Notify(NotificationSeverity.Success, "Saved", "Material outward entry saved successfully.");
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
        var target = !string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/gate/material-outward";
        Navigation.NavigateTo(target);
    }
}

