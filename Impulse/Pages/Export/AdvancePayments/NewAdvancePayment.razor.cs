using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services.Export;
using Impulse.Services.Company;
using Radzen;

namespace Impulse.Pages.Export.AdvancePayments
{
    public partial class NewAdvancePayment
    {
        [Parameter] public int? EntryID { get; set; }

        [Inject] private IAdvancePaymentService AdvancePaymentService { get; set; } = default!;
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

        private AdvancePaymentViewModel Payment { get; set; } = new AdvancePaymentViewModel();

        private List<CustomerLookupModel> Customers { get; set; } = new();
        private CustomerLookupModel? SelectedCustomer { get; set; }

        private List<string> Countries { get; set; } = new();
        private string SelectedCountry { get; set; } = string.Empty;

        private List<string> CustomerOrders { get; set; } = new();
        private List<AdvancePaymentBankModel> Banks { get; set; } = new();

        private string Currency { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();

            if (EntryID.HasValue && EntryID.Value > 0)
            {
                var payment = await AdvancePaymentService.GetAdvancePaymentAsync(EntryID.Value);
                if (payment != null)
                {
                    Payment = payment;
                    SelectedCustomer = Customers.FirstOrDefault(c => c.CustCode == Payment.CustCode);
                    if (SelectedCustomer != null)
                    {
                        await PopulateCountriesForCustomer();
                        SelectedCountry = Payment.Country;
                        await PopulateOrderAndDetails();
                    }
                }
            }
        }

        private async Task LoadLookups()
        {
            try
            {
                var allCustomers = await CustomerOrderService.GetCustomersAsync();
                Customers = allCustomers.OrderBy(c => c.Name).ToList();

                Banks = await AdvancePaymentService.GetBanksAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error", Detail = $"Failed to load lookups: {ex.Message}" });
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

        private async Task HandleCustomerChanged(CustomerLookupModel? customer)
        {
            SelectedCustomer = customer;
            Payment.CustCode = customer?.CustCode ?? string.Empty;

            if (customer != null && customer.CustCode != "0")
            {
                await PopulateCountriesForCustomer();
                if (Countries.Count == 1)
                {
                    SelectedCountry = Countries.First();
                    await HandleCustomerCountryChanged();
                }
                else
                {
                    SelectedCountry = string.Empty;
                }
            }
            else
            {
                Countries.Clear();
                SelectedCountry = string.Empty;
                CustomerOrders.Clear();
                Payment.OrderNo = string.Empty;
                Currency = string.Empty;
                Payment.ImporterAddress = string.Empty;
            }
        }

        private async Task PopulateCountriesForCustomer()
        {
            if (SelectedCustomer != null)
            {
                var custCountries = await CustomerOrderService.GetCountriesForCustomerAsync(SelectedCustomer.CustCode);
                Countries = custCountries.OrderBy(c => c).ToList();
            }
        }

        private async Task HandleCustomerCountryChanged()
        {
            Payment.Country = SelectedCountry;
            await PopulateOrderAndDetails();
        }

        private async Task PopulateOrderAndDetails()
        {
            if (!string.IsNullOrEmpty(Payment.CustCode) && !string.IsNullOrEmpty(Payment.Country))
            {
                CustomerOrders = await AdvancePaymentService.GetOrdersAsync(Payment.CustCode, Payment.Country);
                
                var details = await AdvancePaymentService.GetCustomerDetailsAsync(Payment.CustCode, Payment.Country);
                Currency = details.Currency;
                
                // Only overwrite address if we are creating new or it's currently empty
                if (Payment.EntryID == 0 || string.IsNullOrWhiteSpace(Payment.ImporterAddress))
                {
                    Payment.ImporterAddress = details.Address;
                }
            }
            else
            {
                CustomerOrders.Clear();
                Currency = string.Empty;
            }
        }

        private void OnOrderNoChanged()
        {
            // Optional: additional logic if order selection changes
        }

        private async Task SaveAndClose()
        {
            if (string.IsNullOrWhiteSpace(Payment.CustCode) || string.IsNullOrWhiteSpace(Payment.Country))
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Validation", Detail = "Please select a Customer and Country." });
                return;
            }

            if (Payment.Amount <= 0)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Validation", Detail = "Amount must be greater than zero." });
                return;
            }

            if (string.IsNullOrWhiteSpace(Payment.BankAccNo))
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Validation", Detail = "Please select a Bank." });
                return;
            }

            try
            {
                Payment.UserName = "BlazorUser"; // Replace with actual logged-in user if available
                Payment.MachineName = "BlazorClient";

                bool success = await AdvancePaymentService.SaveAdvancePaymentAsync(Payment);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Success, Summary = "Success", Detail = "Advance Payment saved successfully." });
                    NavigationManager.NavigateTo("/export/advance-payment-list");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Save Failed", Detail = ex.Message });
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo("/export/advance-payment-list");
        }
    }
}
