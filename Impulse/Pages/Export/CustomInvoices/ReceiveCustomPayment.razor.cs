using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Radzen;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;

namespace Impulse.Pages.Export.CustomInvoices
{
    public partial class ReceiveCustomPayment : ComponentBase
    {
        [Parameter]
        public string? CustomInvoiceNo { get; set; }

        [Parameter]
        public string? CustCode { get; set; }

        public bool IsLoading { get; set; } = false;
        public ReceivePaymentModel PaymentModel { get; set; } = new() { IsCustomPayment = true };

        public CustomerLookupModel? SelectedCustomer { get; set; }
        public List<GenericDropDownModel> Banks { get; set; } = new();
        public List<string> CustomerInvoices { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            Banks = await CustomPaymentService.GetBanksAsync();

            if (!string.IsNullOrEmpty(CustCode))
            {
                var customers = await CustomInvoiceService.GetCustomersAsync();
                SelectedCustomer = customers.FirstOrDefault(c => c.CustCode == CustCode);
                if (SelectedCustomer != null)
                {
                    await OnCustomerChanged(SelectedCustomer);
                }
            }

            if (!string.IsNullOrEmpty(CustomInvoiceNo))
            {
                PaymentModel.CustomInvoice = CustomInvoiceNo;
                await OnInvoiceChanged(CustomInvoiceNo);
            }
        }

        public async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            var allCustomers = await CustomInvoiceService.GetCustomersAsync();
            return allCustomers.Where(c => c.CustCode != "HICO" && 
                (string.IsNullOrEmpty(searchText) || 
                 c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                 c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public async Task OnCustomerChanged(CustomerLookupModel cust)
        {
            SelectedCustomer = cust;
            PaymentModel.CustCode = cust.CustCode;
            
            var countries = await CustomInvoiceService.GetCountriesForCustomerAsync(cust.CustCode);
            PaymentModel.Country = countries.FirstOrDefault() ?? "Unknown";
            
            // Assume USD if not found or look up using GetCurrencyForCustomerAsync
            PaymentModel.Currency = await CustomInvoiceService.GetCurrencyForCustomerAsync(cust.CustCode, PaymentModel.Country);

            PaymentModel.ExchangeRate = await CustomPaymentService.GetCurrencyExchangeRateAsync(PaymentModel.Currency);

            CustomerInvoices = await CustomPaymentService.GetCustomInvoicesForPaymentAsync(cust.CustCode);
        }

        public async Task<IEnumerable<string>> SearchInvoices(string searchText)
        {
            await Task.Delay(10);
            if (string.IsNullOrEmpty(searchText))
                return CustomerInvoices;
                
            return CustomerInvoices.Where(x => x.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        public async Task OnInvoiceChanged(string invoice)
        {
            PaymentModel.CustomInvoice = invoice;
            PaymentModel.Amount = await CustomPaymentService.GetCustomInvoiceBalanceAsync(invoice);
        }

        public async Task SavePayment()
        {
            if (string.IsNullOrEmpty(PaymentModel.CustomInvoice))
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Warning", Detail = "Please select a valid invoice." });
                return;
            }

            if (PaymentModel.BankID == 0)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Warning", Detail = "Please select a bank." });
                return;
            }

            if (PaymentModel.Amount <= 0)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Warning", Detail = "Invalid amount entered." });
                return;
            }

            if (PaymentModel.ExchangeRate <= 0)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Warning", Detail = "Invalid exchange rate." });
                return;
            }

            try
            {
                IsLoading = true;
                bool success = await CustomPaymentService.SaveReceivePaymentAsync(PaymentModel);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Success, Summary = "Success", Detail = "Payment received successfully." });
                    NavManager.NavigateTo("/export/custom-payment-status");
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error", Detail = "Failed to save payment." });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error", Detail = ex.Message });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void GoBack()
        {
            NavManager.NavigateTo("/export/custom-payment-status");
        }
    }
}
