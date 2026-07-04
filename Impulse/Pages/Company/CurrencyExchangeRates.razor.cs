using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services.Company;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Pages.Company
{
    public partial class CurrencyExchangeRates : ComponentBase
    {
        [Inject] private ICurrencyExchangeRatesService RatesService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavManager { get; set; } = default!;

        private DateTime SelectedDate { get; set; } = DateTime.Today;
        private List<CurrencyExchangeRateModel> ratesData = new();
        private bool isLoading = true;

        protected override async Task OnInitializedAsync()
        {
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            isLoading = true;
            try
            {
                ratesData = await RatesService.GetRatesForDateAsync(SelectedDate);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load rates: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task RefreshData()
        {
            await LoadDataAsync();
            NotificationService.Notify(NotificationSeverity.Info, "Refreshed", "Data refreshed.");
        }

        private void StartEdit(CurrencyExchangeRateModel item)
        {
            if (item.IsEditing) return;

            // Reset any others
            foreach (var r in ratesData)
                r.IsEditing = false;

            item.EditRate = item.ExchRate;
            item.IsEditing = true;
        }

        private async Task HandleKeyDown(KeyboardEventArgs e, CurrencyExchangeRateModel item)
        {
            if (e.Key == "Enter")
            {
                await SaveRateAsync(item);
            }
            else if (e.Key == "Escape")
            {
                item.IsEditing = false;
            }
        }

        private async Task SaveRateAsync(CurrencyExchangeRateModel item)
        {
            if (!item.IsEditing) return;

            try
            {
                // Only save if it actually changed, or if it's a new entry (EntryID == 0) and user deliberately pressed enter/blur
                if (item.EntryID == 0 || Math.Abs(item.ExchRate - item.EditRate) > 0.0001f)
                {
                    int newId = await RatesService.SaveRateAsync(SelectedDate, item.Currency, item.EditRate, item.EntryID);
                    
                    item.EntryID = newId;
                    item.ExchRate = item.EditRate;
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", $"Rate for {item.Currency} updated.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to save: {ex.Message}");
            }
            finally
            {
                item.IsEditing = false;
            }
        }

        private void NavigateToDashboard()
        {
            NavManager.NavigateTo("/company");
        }
    }
}
