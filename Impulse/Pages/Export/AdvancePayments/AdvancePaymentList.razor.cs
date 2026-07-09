using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services.Export;
using Impulse.Services.Company;
using Radzen;
using Impulse.Services;
using Impulse.Constants;

namespace Impulse.Pages.Export.AdvancePayments
{
    public partial class AdvancePaymentList
    {
        [Inject] private IAdvancePaymentService AdvancePaymentService { get; set; } = default!;
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private IJSRuntime JSRuntime { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;

        private List<AdvancePaymentListViewModel> Payments { get; set; } = new();
        private bool IsLoading { get; set; } = false;

        private List<CustomerLookupModel> Customers { get; set; } = new();
        private CustomerLookupModel? SelectedCustomer { get; set; }
        
        private List<string> Countries { get; set; } = new();
        private string SelectedCountry { get; set; } = "<All Countries>";

        private int SelectedRange { get; set; } = 3; // Default to Last 60 Days (like legacy Case 3)
        private DateTime DtFrom { get; set; }
        private DateTime DtTo { get; set; }
        private bool IsCustomRange => SelectedRange == 4;

        // Inline edit state
        private int EditingExchRateEntryID { get; set; } = 0;
        private decimal NewExchRate { get; set; }

        protected override async Task OnInitializedAsync()
        {
            SetDateRange();
            await LoadLookups();
            await LoadData();
        }

        private async Task LoadLookups()
        {
            try
            {
                var allCustomers = await CustomerOrderService.GetCustomersAsync();
                Customers = allCustomers.OrderBy(c => c.Name).ToList();

                var allCountries = await CustomerOrderService.GetAllCountriesAsync();
                Countries = allCountries.OrderBy(c => c).ToList();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error", Detail = $"Failed to load lookups: {ex.Message}" });
            }
        }

        private void SetDateRange()
        {
            DtTo = DateTime.Today;
            switch (SelectedRange)
            {
                case 0: DtFrom = DateTime.Today; break;
                case 1: DtFrom = DateTime.Today.AddDays(-15); break;
                case 2: DtFrom = DateTime.Today.AddDays(-30); break;
                case 3: DtFrom = DateTime.Today.AddDays(-60); break;
                case 4: /* Custom Range, do not alter DtFrom */ break;
            }
        }

        private async Task OnRangeChanged()
        {
            if (!IsCustomRange)
            {
                SetDateRange();
                await LoadData();
            }
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Customers;
            return await Task.FromResult(Customers.Where(c => 
                (c.Name != null && c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (c.CustCode != null && c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            ).ToList());
        }

        private async Task OnCustomerSelected(CustomerLookupModel? customer)
        {
            SelectedCustomer = customer;
            if (customer != null && customer.CustCode != "0")
            {
                var custCountries = await CustomerOrderService.GetCountriesForCustomerAsync(customer.CustCode);
                Countries = custCountries.OrderBy(c => c).ToList();
                if (Countries.Count == 1)
                {
                    SelectedCountry = Countries.First();
                }
                else
                {
                    SelectedCountry = "<All Countries>";
                }
            }
            else
            {
                var allCountries = await CustomerOrderService.GetAllCountriesAsync();
                Countries = allCountries.OrderBy(c => c).ToList();
                SelectedCountry = "<All Countries>";
            }
            await LoadData();
        }

        private async Task LoadData()
        {
            IsLoading = true;
            try
            {
                string custCode = SelectedCustomer?.CustCode ?? "0";
                Payments = await AdvancePaymentService.GetAdvancePaymentListAsync(DtFrom, DtTo, custCode, SelectedCountry);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error", Detail = $"Failed to load data: {ex.Message}" });
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void AddNew()
        {
            NavigationManager.NavigateTo("/export/advance-payment/new");
        }

        private void Edit(AdvancePaymentListViewModel item)
        {
            if (item.Status == "Posted to Financial")
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Cannot Edit", Detail = "Already Posted, Can't Edit." });
                return;
            }
            NavigationManager.NavigateTo($"/export/advance-payment/{item.EntryID}");
        }

        private async Task Print(AdvancePaymentListViewModel item)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "rptAdvancePayments.rpt",
                    SelectionFormula = $"{{FCustAdvancePayments.EntryID}}={item.EntryID}"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Print Error", Detail = ex.Message });
            }
        }

        private async Task BeginEditExchangeRate(AdvancePaymentListViewModel item)
        {
            EditingExchRateEntryID = item.EntryID;
            NewExchRate = item.ExchRate;
            
            // Focus input using JSInterop
            await Task.Delay(100);
            try { await JSRuntime.InvokeVoidAsync("eval", $"document.getElementById('exch_{item.EntryID}')?.focus()"); } catch { }
        }

        private async Task HandleExchRateKeyUp(KeyboardEventArgs e, AdvancePaymentListViewModel item)
        {
            if (e.Key == "Enter")
            {
                await SaveExchangeRate(item);
            }
            else if (e.Key == "Escape")
            {
                EditingExchRateEntryID = 0;
            }
        }

        private async Task SaveExchangeRate(AdvancePaymentListViewModel item)
        {
            if (EditingExchRateEntryID == item.EntryID)
            {
                if (NewExchRate <= 0)
                {
                    NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Invalid Rate", Detail = "Invalid Exch. Rate" });
                    return;
                }

                try
                {
                    bool success = await AdvancePaymentService.UpdateExchangeRateAsync(item.EntryID, NewExchRate);
                    if (success)
                    {
                        item.ExchRate = NewExchRate;
                        EditingExchRateEntryID = 0;
                        // Status will auto-compute to "Ready to Post" since it's a getter, but UI needs re-render
                    }
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Save Failed", Detail = ex.Message });
                }
            }
        }

        private void PostToFinancial(AdvancePaymentListViewModel item)
        {
            if (string.IsNullOrWhiteSpace(item.AdviceNo))
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Required", Detail = "Please Enter Advice No. (Edit the voucher to add it)." });
                return;
            }

            NavigationManager.NavigateTo($"/export/advance-payment-posting/{item.EntryID}");
        }

        private bool ShowUsageModal { get; set; }
        private List<AdvancePaymentUsageViewModel>? UsageDetails { get; set; }

        private async Task ShowUsageDetail(AdvancePaymentListViewModel item)
        {
            ShowUsageModal = true;
            UsageDetails = null; // show loading
            StateHasChanged();

            try
            {
                UsageDetails = await AdvancePaymentService.GetUsageDetailsAsync(item.EntryID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error", Detail = ex.Message });
                UsageDetails = new List<AdvancePaymentUsageViewModel>();
            }
        }

        private void CloseUsageModal()
        {
            ShowUsageModal = false;
        }
    }
}
