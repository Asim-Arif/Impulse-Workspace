using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;

namespace Impulse.Pages.Export.Invoices
{
    public partial class PackingLabels : ComponentBase
    {
        [Inject] private IPackingLabelsService PackingLabelsService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = true;

        // Form bindings & lists
        private List<CustomerLookupModel> customers = new();
        private CustomerLookupModel? selectedCustomer;
        private List<string> countries = new();
        private string selectedCountry = string.Empty;
        private string currency = string.Empty;
        private List<string> invoices = new();
        private string selectedInvoice = string.Empty;
        private int totalCartons = 0;
        private int selectedFromCarton = 1;
        private int selectedToCarton = 1;
        private bool printInner = false;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                customers = await PackingLabelsService.GetActiveCustomersAsync();
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
            countries.Clear();
            currency = string.Empty;
            invoices.Clear();
            selectedInvoice = string.Empty;
            totalCartons = 0;
            selectedFromCarton = 1;
            selectedToCarton = 1;

            if (selectedCustomer != null)
            {
                try
                {
                    countries = await PackingLabelsService.GetCountriesForCustomerAsync(selectedCustomer.CustCode);
                    if (countries.Any())
                    {
                        selectedCountry = countries.First();
                        await OnCountryChangedInternal(selectedCountry);
                    }
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Data Error", "Failed to load countries: " + ex.Message);
                }
            }
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? string.Empty;
            await OnCountryChangedInternal(selectedCountry);
        }

        private async Task OnCountryChangedInternal(string country)
        {
            currency = string.Empty;
            invoices.Clear();
            selectedInvoice = string.Empty;
            totalCartons = 0;
            selectedFromCarton = 1;
            selectedToCarton = 1;

            if (selectedCustomer != null && !string.IsNullOrEmpty(country))
            {
                try
                {
                    var settings = await PackingLabelsService.GetLabelSettingsAsync(selectedCustomer.CustCode, country);
                    if (settings != null)
                    {
                        currency = settings.Curr;
                    }

                    invoices = await PackingLabelsService.GetCustomInvoicesAsync(selectedCustomer.CustCode, country);
                    if (invoices.Any())
                    {
                        selectedInvoice = invoices.First();
                        await OnInvoiceChangedInternal(selectedInvoice);
                    }
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Data Error", "Failed to load details: " + ex.Message);
                }
            }
        }

        private async Task OnInvoiceChanged(ChangeEventArgs e)
        {
            selectedInvoice = e.Value?.ToString() ?? string.Empty;
            await OnInvoiceChangedInternal(selectedInvoice);
        }

        private async Task OnInvoiceChangedInternal(string invoice)
        {
            totalCartons = 0;
            selectedFromCarton = 1;
            selectedToCarton = 1;

            if (!string.IsNullOrEmpty(invoice))
            {
                try
                {
                    totalCartons = await PackingLabelsService.GetTotalCartonsAsync(invoice);
                    if (totalCartons > 0)
                    {
                        selectedFromCarton = 1;
                        selectedToCarton = totalCartons;
                    }
                    else
                    {
                        NotificationServiceManager.ShowInformation("Information", "Packing List Not Generated For The Selected Invoice.");
                    }
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Data Error", "Failed to load carton total: " + ex.Message);
                }
            }
        }

        private async Task OpenOptionsBarMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private async Task OnPrintClick()
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry) || string.IsNullOrEmpty(selectedInvoice))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select all required fields.");
                return;
            }

            if (selectedFromCarton <= 0 || selectedToCarton <= 0)
            {
                NotificationServiceManager.ShowWarning("Validation", "Invalid Carton No. Selected.");
                return;
            }

            try
            {
                // Generate print mappings
                await PackingLabelsService.GeneratePrintPackingLabelsAsync(selectedInvoice);

                // Fetch labels configuration
                var settings = await PackingLabelsService.GetLabelSettingsAsync(selectedCustomer.CustCode, selectedCountry);
                string reportName = printInner ? settings?.InnerPackingLabel : settings?.OuterPackingLabel;

                if (string.IsNullOrEmpty(reportName))
                {
                    NotificationServiceManager.ShowWarning("Config Error", $"No {(printInner ? "inner" : "outer")} packing label configured for this customer and country.");
                    return;
                }

                var request = new ReportRequest
                {
                    ReportName = reportName,
                    SelectionFormula = $"{{VrptProformas.CustomInvoice}}='{selectedInvoice}' AND {{CustomPList.CartonFrom}} In ({selectedFromCarton} to {selectedToCarton})"
                };

                await ReportNavigation.PrintReportAsync(request);
                NotificationServiceManager.ShowSuccess("Success", "Report dispatched successfully.");
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", "Failed to generate packing labels: " + ex.Message);
            }
        }

        private async Task PrintItemLabel(ItemClickEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedInvoice))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select an Invoice first.");
                return;
            }

            try
            {
                var request = new ReportRequest
                {
                    ReportName = "ItemLabel.rpt",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@InvoiceNo", selectedInvoice }
                    }
                };

                await ReportNavigation.PrintReportAsync(request);
                NotificationServiceManager.ShowSuccess("Success", "Item Label report dispatched successfully.");
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", "Failed to print item label: " + ex.Message);
            }
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo("/export");
        }
    }
}
