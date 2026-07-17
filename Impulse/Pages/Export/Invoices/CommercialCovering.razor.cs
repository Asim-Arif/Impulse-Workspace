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
    public partial class CommercialCovering : ComponentBase
    {
        [Inject] private ICommercialCoveringService CommercialCoveringService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = true;
        private FormState formState = FormState.View;

        // Dropdowns & Selections
        private List<CustomerLookupModel> customers = new();
        private CustomerLookupModel? selectedCustomer;
        private List<CustomerCountryModel> countries = new();
        private string selectedCountry = string.Empty;
        private List<string> invoices = new();
        private string selectedInvoice = string.Empty;

        // Table Rows List
        private List<CoveringLetterItemViewModel> items = new();
        private List<CoveringLetterItemViewModel> originalItems = new();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                customers = await CommercialCoveringService.GetActiveCustomersAsync();
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
            invoices.Clear();
            selectedInvoice = string.Empty;
            items.Clear();

            if (selectedCustomer != null)
            {
                try
                {
                    countries = await CommercialCoveringService.GetCountriesForCustomerAsync(selectedCustomer.CustCode);
                    if (countries.Count > 0)
                    {
                        selectedCountry = countries[0].Country;
                        await LoadInvoicesAsync();
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
            invoices.Clear();
            selectedInvoice = string.Empty;
            items.Clear();

            if (!string.IsNullOrEmpty(selectedCountry))
            {
                await LoadInvoicesAsync();
            }
        }

        private async Task OnInvoiceChanged(ChangeEventArgs e)
        {
            selectedInvoice = e.Value?.ToString() ?? string.Empty;
            items.Clear();

            if (!string.IsNullOrEmpty(selectedInvoice))
            {
                if (formState == FormState.View)
                {
                    await LoadLetterItemsAsync(selectedInvoice);
                }
                else if (formState == FormState.Add)
                {
                    PrepopulateDefaultItems();
                }
            }
        }

        private async Task LoadInvoicesAsync()
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry))
                return;

            try
            {
                if (formState == FormState.View)
                {
                    invoices = await CommercialCoveringService.GetInvoicesWithLettersAsync(selectedCustomer.CustCode, selectedCountry);
                }
                else
                {
                    invoices = await CommercialCoveringService.GetInvoicesWithoutLettersAsync(selectedCustomer.CustCode, selectedCountry);
                }

                if (invoices.Count > 0)
                {
                    selectedInvoice = invoices[0];
                    if (formState == FormState.View)
                    {
                        await LoadLetterItemsAsync(selectedInvoice);
                    }
                    else if (formState == FormState.Add)
                    {
                        PrepopulateDefaultItems();
                    }
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

        private async Task LoadLetterItemsAsync(string customInvoice)
        {
            try
            {
                items = await CommercialCoveringService.GetCoveringLetterItemsAsync(customInvoice);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Data Error", "Failed to load covering letter details: " + ex.Message);
            }
        }

        private void PrepopulateDefaultItems()
        {
            items = new List<CoveringLetterItemViewModel>
            {
                new CoveringLetterItemViewModel { EntryText = "Commercial Invoice", EntryDetail = "(3)" },
                new CoveringLetterItemViewModel { EntryText = "PACKING LIST", EntryDetail = "(1)" },
                new CoveringLetterItemViewModel { EntryText = "GSP FORM (A)", EntryDetail = "(1)" },
                new CoveringLetterItemViewModel { EntryText = "Bill of lading", EntryDetail = "(3)" },
                new CoveringLetterItemViewModel { EntryText = "FDA Package", EntryDetail = "(2)" },
                new CoveringLetterItemViewModel { EntryText = "Letter to Customer", EntryDetail = "(1)" },
                new CoveringLetterItemViewModel { EntryText = "Any Instruction Letter", EntryDetail = "(1)" },
                new CoveringLetterItemViewModel { EntryText = "AWB/EMS/Courier Receipts", EntryDetail = "(1)" }
            };
        }

        private void AddRow()
        {
            items.Add(new CoveringLetterItemViewModel { EntryText = "", EntryDetail = "" });
        }

        private void RemoveRow(int index)
        {
            if (index >= 0 && index < items.Count)
            {
                items.RemoveAt(index);
            }
        }

        private async Task StartAddNew()
        {
            formState = FormState.Add;
            invoices.Clear();
            selectedInvoice = string.Empty;
            items.Clear();
            await LoadInvoicesAsync();
        }

        private void StartEdit()
        {
            if (string.IsNullOrEmpty(selectedInvoice))
                return;

            formState = FormState.Edit;
            originalItems = items.Select(i => new CoveringLetterItemViewModel { 
                EntryText = i.EntryText, 
                EntryDetail = i.EntryDetail 
            }).ToList();
        }

        private async Task CancelChanges()
        {
            if (formState == FormState.Add)
            {
                formState = FormState.View;
                await LoadInvoicesAsync();
            }
            else if (formState == FormState.Edit)
            {
                formState = FormState.View;
                items = originalItems;
            }
        }

        private async Task SaveLetter()
        {
            if (string.IsNullOrEmpty(selectedInvoice))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select an invoice.");
                return;
            }

            if (!items.Any(i => !string.IsNullOrWhiteSpace(i.EntryText)))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please enter at least one covering letter item description.");
                return;
            }

            try
            {
                bool success = await CommercialCoveringService.SaveCoveringLetterItemsAsync(selectedInvoice, items);
                if (success)
                {
                    NotificationServiceManager.ShowSuccess("Success", "Commercial covering letter details saved successfully.");
                    formState = FormState.View;
                    await LoadInvoicesAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Save Error", "Failed to save covering letter: " + ex.Message);
            }
        }

        private async Task PrintLetter()
        {
            if (string.IsNullOrEmpty(selectedInvoice))
                return;

            try
            {
                var request = new ReportRequest
                {
                    ReportName = "rptCommercialCoveringLetter.rpt",
                    SelectionFormula = $"{{VCommercialCoveringLetter.CustomInvoice}}='{selectedInvoice}'"
                };

                await ReportNavigation.PrintReportAsync(request);
                NotificationServiceManager.ShowSuccess("Success", "Covering letter report opened.");
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", "Failed to print covering letter: " + ex.Message);
            }
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo("/export");
        }
    }
}
