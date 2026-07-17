using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Export.Invoices
{
    public enum FormState
    {
        View,
        Add,
        Edit
    }

    public partial class ShippingInstructions : ComponentBase
    {
        [Inject] private IShippingInstructionsService ShippingInstructionsService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;

        private bool isLoading = true;
        private FormState formState = FormState.View;

        // Dropdown options
        private List<CustomerLookupModel> customers = new();
        private CustomerLookupModel? selectedCustomer;
        private List<string> countries = new();
        private string selectedCountry = string.Empty;
        private List<string> invoices = new();
        private string selectedInvoice = string.Empty;

        // Form Model
        private ShippingInstructionsViewModel model = new();
        
        // Cache to support Cancel on Edit
        private ShippingInstructionsViewModel originalModel = new();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                customers = await ShippingInstructionsService.GetActiveCustomersAsync();
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
            ClearForm();

            if (selectedCustomer != null)
            {
                try
                {
                    countries = await ShippingInstructionsService.GetCountriesForCustomerAsync(selectedCustomer.CustCode);
                    if (countries.Count > 0)
                    {
                        selectedCountry = countries[0];
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
            ClearForm();

            if (!string.IsNullOrEmpty(selectedCountry))
            {
                await LoadInvoicesAsync();
            }
        }

        private async Task OnInvoiceChanged(ChangeEventArgs e)
        {
            selectedInvoice = e.Value?.ToString() ?? string.Empty;
            ClearForm();

            if (!string.IsNullOrEmpty(selectedInvoice))
            {
                if (formState == FormState.View)
                {
                    await LoadInstructionsForInvoiceAsync(selectedInvoice);
                }
                else if (formState == FormState.Add)
                {
                    await PrepopulateDefaultInstructionsAsync(selectedInvoice);
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
                    invoices = await ShippingInstructionsService.GetInvoicesWithInstructionsAsync(selectedCustomer.CustCode, selectedCountry);
                }
                else
                {
                    invoices = await ShippingInstructionsService.GetInvoicesWithoutInstructionsAsync(selectedCustomer.CustCode, selectedCountry);
                }

                if (invoices.Count > 0)
                {
                    selectedInvoice = invoices[0];
                    if (formState == FormState.View)
                    {
                        await LoadInstructionsForInvoiceAsync(selectedInvoice);
                    }
                    else if (formState == FormState.Add)
                    {
                        await PrepopulateDefaultInstructionsAsync(selectedInvoice);
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

        private async Task LoadInstructionsForInvoiceAsync(string customInvoice)
        {
            try
            {
                var data = await ShippingInstructionsService.GetShippingInstructionsAsync(customInvoice);
                if (data != null)
                {
                    model = data;
                }
                else
                {
                    ClearForm();
                    model.CustomInvoice = customInvoice;
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Data Error", "Failed to load shipping instructions: " + ex.Message);
            }
        }

        private async Task PrepopulateDefaultInstructionsAsync(string customInvoice)
        {
            try
            {
                ClearForm();
                model.CustomInvoice = customInvoice;

                var (totalCartons, invoiceOf) = await ShippingInstructionsService.GetTotalCartonsAndInvoiceOfAsync(customInvoice);

                model.GoodsDesc = string.IsNullOrEmpty(invoiceOf) ? "SURGICAL INSTRUMENTS." : invoiceOf;
                model.Encloseres = "1.   INVOICES 5 COPIES.\n2.   PACKING LIST 2 COPIES.\n3.   VALUATION FORM (A).\n4.   FORM E ONE SET.\n5.   UNDERTAKING\n6.   COMMERCIAL EXPORTER DECLARATION";

                // Substring logic: Left(cmbInvNoNew, InStr(5, cmbInvNoNew, "-") + 4)
                int index = customInvoice.IndexOf('-', 5);
                string prefix = index != -1 && index + 4 < customInvoice.Length
                    ? customInvoice.Substring(0, index + 5)
                    : customInvoice;

                if (totalCartons > 0)
                {
                    model.MarksnNumbers = $"{prefix}/1 TO {totalCartons}";
                    model.NonKind = $"({totalCartons}) CARTONS\nCONTAINING DISPOSABLE\nSURGICAL INSTRUMENTS";
                }
                else
                {
                    model.MarksnNumbers = $"{prefix}/1 TO ";
                    model.NonKind = " CARTONS\nCONTAINING DISPOSABLE\nSURGICAL INSTRUMENTS";
                }

                if (totalCartons > 0)
                {
                    double measurementVal = await ShippingInstructionsService.GetDefaultBoxSizeMeasurementAsync(totalCartons);
                    model.Measurement = measurementVal.ToString();
                }
                else
                {
                    model.Measurement = "0";
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Data Error", "Failed to populate defaults: " + ex.Message);
            }
        }

        private void ClearForm()
        {
            model = new ShippingInstructionsViewModel();
        }

        private async Task StartAddNew()
        {
            formState = FormState.Add;
            invoices.Clear();
            selectedInvoice = string.Empty;
            ClearForm();
            await LoadInvoicesAsync();
        }

        private void StartEdit()
        {
            if (string.IsNullOrEmpty(selectedInvoice))
                return;

            formState = FormState.Edit;
            originalModel = new ShippingInstructionsViewModel
            {
                CustomInvoice = model.CustomInvoice,
                GoodsDesc = model.GoodsDesc,
                GWeight = model.GWeight,
                Measurement = model.Measurement,
                MarksnNumbers = model.MarksnNumbers,
                NonKind = model.NonKind,
                Encloseres = model.Encloseres,
                SpecialInstructions = model.SpecialInstructions
            };
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
                model = originalModel;
            }
        }

        private async Task SaveInstructions()
        {
            if (string.IsNullOrEmpty(selectedInvoice))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select an invoice.");
                return;
            }

            try
            {
                model.CustomInvoice = selectedInvoice;
                bool success = await ShippingInstructionsService.SaveShippingInstructionsAsync(model);
                if (success)
                {
                    NotificationServiceManager.ShowSuccess("Success", "Shipping instructions saved successfully.");
                    formState = FormState.View;
                    await LoadInvoicesAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Save Error", "Failed to save instructions: " + ex.Message);
            }
        }

        private async Task PrintWithValue()
        {
            await PrintInstructionsAsync(false);
        }

        private async Task PrintWithoutValue()
        {
            await PrintInstructionsAsync(true);
        }

        private async Task PrintInstructionsAsync(bool hideValue)
        {
            if (string.IsNullOrEmpty(selectedInvoice))
                return;

            try
            {
                var request = new ReportRequest
                {
                    ReportName = "rptShipping.rpt",
                    SelectionFormula = $"{{VShippingInstructions.CustomInvoice}}='{selectedInvoice}'",
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "HideValue", hideValue ? "True" : "False" }
                    }
                };

                await ReportNavigation.PrintReportAsync(request);
                NotificationServiceManager.ShowSuccess("Success", "Shipping instructions report opened.");
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

        private async Task OpenPrintMenu(Microsoft.AspNetCore.Components.Web.MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("printOptionsMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }
    }
}
