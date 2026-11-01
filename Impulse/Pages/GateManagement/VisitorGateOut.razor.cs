using System;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.GateManagement;

public partial class VisitorGateOut : ComponentBase
{
    [Inject] private IMiscGateInDataAccess DataAccess { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private NotificationService NotificationService { get; set; } = default!;

    [Parameter]
    [SupplyParameterFromQuery(Name = "entryId")]
    public int? EntryId { get; set; }

    protected MiscGateOutModel Model { get; set; } = new();
    protected string SearchQuery { get; set; } = "";
    protected bool IsLoading { get; set; }
    protected bool IsSaving { get; set; }

    protected string TimeOutInputValue => Model.TimeOut.ToString("yyyy-MM-ddTHH:mm");
    protected string StayDurationText
    {
        get
        {
            var dur = Model.StayDuration;
            if (dur.TotalMinutes < 1) return "Just arrived";
            if (dur.TotalHours < 1) return $"{(int)dur.TotalMinutes} mins";
            return $"{(int)dur.TotalHours}h {dur.Minutes}m";
        }
    }

    protected override async Task OnInitializedAsync()
    {
        if (EntryId.HasValue && EntryId.Value > 0)
        {
            await LoadVisitorAsync(EntryId.Value);
        }
    }

    protected override async Task OnParametersSetAsync()
    {
        if (EntryId.HasValue && EntryId.Value > 0 && EntryId.Value != Model.EntryId)
        {
            await LoadVisitorAsync(EntryId.Value);
        }
    }

    protected async Task LoadVisitorAsync(int id)
    {
        IsLoading = true;
        try
        {
            var outModel = await DataAccess.GetVisitorGateOutAsync(id);
            if (outModel == null)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Not Found", "Visitor record was not found.");
                return;
            }

            Model = outModel;
            Model.VerificationInput = Model.CardNo;
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

    protected void OnTimeOutChanged(ChangeEventArgs args)
    {
        if (DateTime.TryParse(args.Value?.ToString(), out var dt))
        {
            Model.TimeOut = dt;
        }
    }

    protected async Task OnSearchKeyUp(KeyboardEventArgs e)
    {
        if (e.Key == "Enter")
        {
            await SearchVisitorAsync();
        }
    }

    protected async Task SearchVisitorAsync()
    {
        if (string.IsNullOrWhiteSpace(SearchQuery))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Notice", "Please enter a badge card number or visit code.");
            return;
        }

        var query = SearchQuery.Trim();
        int foundId = 0;
        if (int.TryParse(query, out var parsedInt))
        {
            foundId = parsedInt;
        }
        else if (query.StartsWith("CARD-", StringComparison.OrdinalIgnoreCase) || query.StartsWith("VST-", StringComparison.OrdinalIgnoreCase))
        {
            var numericPart = query.Split('-').LastOrDefault();
            if (int.TryParse(numericPart, out var parsed))
            {
                foundId = parsed;
            }
        }

        if (foundId > 0)
        {
            await LoadVisitorAsync(foundId);
        }
        else
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Not Found", $"Could not find a visitor with code '{query}'.");
        }
    }

    protected async Task SaveAsync()
    {
        if (Model.EntryId <= 0) return;

        if (string.IsNullOrWhiteSpace(Model.VerificationInput))
        {
            NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Scan or enter badge card no. before checkout.");
            return;
        }

        IsSaving = true;
        try
        {
            await DataAccess.SaveVisitorGateOutAsync(Model.EntryId, Model.TimeOut, Model.VerificationInput, Environment.UserName);
            NotificationService.Notify(NotificationSeverity.Success, "Gate Out Recorded", $"Visitor '{Model.VisitorName}' successfully checked out.");
            Navigation.NavigateTo("/gate/visitors");
        }
        catch (Exception ex)
        {
            NotificationService.Notify(NotificationSeverity.Error, "Checkout Error", ex.Message);
        }
        finally
        {
            IsSaving = false;
        }
    }

    protected void Cancel() => Navigation.NavigateTo("/gate/visitors");
}

