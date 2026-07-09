using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Export;

namespace Impulse.Pages.Export.Quotations
{
    public partial class CustomerQuotationList : ComponentBase
    {
        [Inject] private ICustomerQuotationService CustomerQuotationService { get; set; } = null!;
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = null!; // For Lookups
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;

        private bool isLoading = true;
        private string searchText = string.Empty;

        // Filters state
        private DateTime dtFrom = DateTime.Today.AddDays(-180);
        private DateTime dtTo = DateTime.Today;
        private string selectedDateRange = "5"; // 5 = Last 180 Days
        private bool isCustomDateRange = false;

        // Lookups
        private List<CustomerLookupModel> customers = new();
        private List<string> countries = new();

        private CustomerLookupModel? selectedCustomerFilter = null;
        private string selectedCountryFilter = "<All Countries>";

        // List records
        private List<CustomerQuotationListItemModel> allQuotations = new();

        // Selected row state
        private CustomerQuotationListItemModel? highlightedQuotation = null;

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();
            await RefreshList();
        }

        private async Task LoadLookups()
        {
            try
            {
                customers = await CustomerOrderService.GetCustomersAsync();
                await LoadCountries();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Loading Lookups", ex.Message);
            }
        }

        private async Task LoadCountries()
        {
            try
            {
                if (selectedCustomerFilter != null && selectedCustomerFilter.CustCode != "0")
                {
                    countries = await CustomerOrderService.GetCountriesForCustomerAsync(selectedCustomerFilter.CustCode);
                }
                else
                {
                    countries = await CustomerOrderService.GetAllCountriesAsync();
                }
                selectedCountryFilter = "<All Countries>";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading countries: {ex.Message}");
            }
        }

        private async Task RefreshList()
        {
            isLoading = true;
            try
            {
                string custCode = selectedCustomerFilter?.CustCode ?? "0";

                allQuotations = await CustomerQuotationService.GetQuotationListAsync(
                    dtFrom,
                    dtTo,
                    custCode,
                    selectedCountryFilter
                );

                highlightedQuotation = null;
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Refreshing List", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private IEnumerable<CustomerQuotationListItemModel> GetFilteredQuotations()
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return allQuotations;
            }
            return allQuotations.Where(q => 
                q.QuotationNo.ToString().Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                (q.CustomerName != null && q.CustomerName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (q.CustCode != null && q.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (q.Country != null && q.Country.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        // ── Blazored Typeahead Filter Event Triggers ──

        public async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return customers;
            return await Task.FromResult(customers.Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            selectedCustomerFilter = cust;
            await LoadCountries();
            await RefreshList();
        }

        public async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountryFilter = e.Value?.ToString() ?? "<All Countries>";
            await RefreshList();
        }

        public async Task OnDateRangeChanged(ChangeEventArgs e)
        {
            selectedDateRange = e.Value?.ToString() ?? "6";
            isCustomDateRange = selectedDateRange == "6";

            dtTo = DateTime.Today;
            switch (selectedDateRange)
            {
                case "0": dtFrom = DateTime.Today; break;
                case "1": dtFrom = DateTime.Today.AddDays(-15); break;
                case "2": dtFrom = DateTime.Today.AddDays(-30); break;
                case "3": dtFrom = DateTime.Today.AddDays(-60); break;
                case "4": dtFrom = DateTime.Today.AddDays(-90); break;
                case "5": dtFrom = DateTime.Today.AddDays(-180); break;
            }
            if (!isCustomDateRange) {
                await RefreshList();
            }
        }

        private void HighlightQuotation(CustomerQuotationListItemModel quotation)
        {
            highlightedQuotation = quotation;
        }

        private async Task OpenRowMenu(MouseEventArgs e, CustomerQuotationListItemModel quotation)
        {
            highlightedQuotation = quotation;
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY, quotation);
        }

        private async Task OpenOptionsBarMenu(MouseEventArgs e)
        {
            // Set vertical offset for top options bar menu to open downwards
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        // ── Row Context Menu Commands ──

        private void AddQuotation()
        {
            NavigationManager.NavigateTo($"/export/new-customer-quotation");
        }

        private void EditQuotation(int quotationNo)
        {
            NavigationManager.NavigateTo($"/export/new-customer-quotation/{quotationNo}");
        }

        private async Task DeleteQuotation(int quotationNo)
        {
            try
            {
                bool deleted = await CustomerQuotationService.DeleteQuotationAsync(quotationNo);
                if (deleted)
                {
                    NotificationServiceManager.ShowSuccess("Deleted", $"Quotation {quotationNo} was deleted successfully.");
                    await RefreshList();
                }
                else
                {
                    NotificationServiceManager.ShowError("Error", $"Failed to delete Quotation {quotationNo}.");
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
        }

        // ── Crystal Reports Triggering ──

        private async Task PrintQuotation(int quotationNo)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "CustomerQuotation.rpt", // VB6 rpt name equivalent
                    SelectionFormula = $"{{CustomerQuotations.QuotationNo}}={quotationNo}"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Report Failed", ex.Message);
            }
        }

        private async Task PrintQuotationDiscount(int quotationNo)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "CustomerQuotationDiscount.rpt", // VB6 rpt name equivalent
                    SelectionFormula = $"{{CustomerQuotations.QuotationNo}}={quotationNo}"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Report Failed", ex.Message);
            }
        }

        private async Task PrintList()
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "CustomerQuotationList.rpt", // General List rpt
                    SelectionFormula = string.Empty
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Report Failed", ex.Message);
            }
        }
    }
}
