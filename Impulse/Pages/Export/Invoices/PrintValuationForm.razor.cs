using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Export.Invoices
{
    public partial class PrintValuationForm : ComponentBase
    {
        [Inject] private IPrintValuationService PrintValuationService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = true;

        // Form bindings & dropdowns
        private List<CustomerLookupModel> customers = new();
        private CustomerLookupModel? selectedCustomer;
        private List<CustomerCountryModel> countries = new();
        private string selectedCountry = string.Empty;
        private string currency = string.Empty;
        private List<string> invoices = new();
        private string selectedInvoice = string.Empty;

        // Option & Check bindings
        private bool printFrontSide = true;
        private bool printValue = true; // Defaults to checked in legacy VB6

        protected override async Task OnInitializedAsync()
        {
            try
            {
                customers = await PrintValuationService.GetCustomersAsync();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Load Error", "Failed to load customers: " + ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            await Task.Yield();
            if (string.IsNullOrWhiteSpace(searchText))
                return customers;
            return customers.Where(c => c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            selectedCustomer = cust;
            selectedCountry = string.Empty;
            currency = string.Empty;
            countries.Clear();
            invoices.Clear();
            selectedInvoice = string.Empty;

            if (selectedCustomer != null)
            {
                try
                {
                    countries = await PrintValuationService.GetCountriesForCustomerAsync(selectedCustomer.CustCode);
                    if (countries.Count > 0)
                    {
                        var firstCountry = countries[0];
                        selectedCountry = firstCountry.Country;
                        currency = firstCountry.Curr;
                        await LoadInvoicesAsync();
                    }
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Data Error", "Failed to load customer settings: " + ex.Message);
                }
            }
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? string.Empty;
            currency = string.Empty;
            invoices.Clear();
            selectedInvoice = string.Empty;

            if (!string.IsNullOrEmpty(selectedCountry))
            {
                var matched = countries.FirstOrDefault(c => c.Country == selectedCountry);
                if (matched != null)
                {
                    currency = matched.Curr;
                }
                await LoadInvoicesAsync();
            }
        }

        private void OnInvoiceChanged(ChangeEventArgs e)
        {
            selectedInvoice = e.Value?.ToString() ?? string.Empty;
        }

        private async Task LoadInvoicesAsync()
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry))
                return;

            try
            {
                invoices = await PrintValuationService.GetInvoicesAsync(selectedCustomer.CustCode, selectedCountry);
                if (invoices.Count > 0)
                {
                    selectedInvoice = invoices[0];
                }
                else
                {
                    selectedInvoice = string.Empty;
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Data Error", "Failed to load invoices: " + ex.Message);
            }
        }

        private void SetPrintSide(bool isFront)
        {
            printFrontSide = isFront;
        }

        private async Task OnPrintClick()
        {
            if (string.IsNullOrEmpty(selectedInvoice))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select an invoice.");
                return;
            }

            try
            {
                string reportName = printFrontSide ? "rptValuationFormA_1.rpt" : "rptValuationFormA_2.rpt";

                var request = new ReportRequest
                {
                    ReportName = reportName,
                    SelectionFormula = $"{{VrptValuationFormA.CustomInvoice}}='{selectedInvoice}'",
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "HideValue", !printValue } // Since company is "IAA", we always pass printValue formula state
                    }
                };

                await ReportNavigation.PrintReportAsync(request);
                NotificationServiceManager.ShowSuccess("Success", "Valuation form report dispatched successfully.");
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", "Failed to print: " + ex.Message);
            }
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo("/export");
        }
    }
}
