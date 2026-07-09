using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Components.Export.NewProformaInvoiceTabs;
using Impulse.Services.Export;
using Impulse.Utilities;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Export.Proformas
{
    public partial class NewProformaInvoice
    {
        [Parameter] public string? PInvoice { get; set; }

        [Inject] private IProformaInvoiceService ProformaService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

        // ── Main Model ─────────────────────────────────────────────────────
        public ProformaInvoiceViewModel Proforma { get; set; } = new();
        public List<ProformaLineItemModel> LineItems { get; set; } = new();
        
        // ── State ──────────────────────────────────────────────────────────
        private bool IsEditMode => !string.IsNullOrEmpty(PInvoice);
        private bool IsLoading = false;
        private int ActiveTab = 1;

        // ── Lookups ────────────────────────────────────────────────────────
        private List<GenericDropDownModel> Companies = new();
        private List<CustomerLookupModel> AllCustomers = new();
        private CustomerLookupModel? SelectedCustomer;
        private List<string> Countries = new();
        private List<GenericDropDownModel> TradeTermsList = new();
        private List<GenericDropDownModel> PaymentTermsList = new();
        private List<GenericDropDownModel> ShippingMethods = new();
        
        // ── Tab references and data ────────────────────────────────────────
        private DetailsTab? DetailsTabRef;
        private OtherInfoTab? OtherInfoTabRef;
        private NewItemsTab? NewItemsTabRef;

        private List<AgentLookupModel> ShippingAgents = new();
        private List<AgentLookupModel> InsuranceAgents = new();
        private List<CustomerBankModel> CustomerBanks = new();
        private List<CompanyBankAccountModel> CompanyBankAccounts = new();
        private List<ShippingAddressPresetModel> ShippingAddressPresets = new();
        private List<InvoiceToPresetModel> InvoiceToPresets = new();

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                await LoadInitialLookups();

                if (IsEditMode)
                {
                    await LoadProforma();
                }
                else
                {
                    Proforma.CompanyRefID = int.TryParse(Companies.FirstOrDefault()?.DropDownValue_ID, out int compId) ? compId : 1;
                    Proforma.PInvoice = await ProformaService.GetNextProformaNumberAsync();
                    await OnCompanyChanged();
                }
            }
            catch (Exception ex)
            {
                ShowError("Error loading page: " + ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task LoadInitialLookups()
        {
            Companies = await ProformaService.GetCompaniesAsync();
            AllCustomers = await ProformaService.GetCustomersAsync();
            TradeTermsList = await ProformaService.GetTradeTermsAsync();
            PaymentTermsList = await ProformaService.GetPaymentTermsAsync();
            ShippingMethods = await ProformaService.GetShippingMethodsAsync();
            ShippingAgents = await ProformaService.GetShippingAgentsAsync();
            InsuranceAgents = await ProformaService.GetInsuranceAgentsAsync();
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            await Task.Delay(1); // Yield
            if (string.IsNullOrWhiteSpace(searchText))
                return AllCustomers;
            return AllCustomers.Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task LoadProforma()
        {
            var p = await ProformaService.GetProformaInvoiceAsync(PInvoice!);
            if (p != null)
            {
                Proforma = p;
                LineItems = p.LineItems;
                
                Proforma.TotalAmount = Math.Round(LineItems.Sum(i => i.Amount), 4);
                Proforma.GrandTotal = Math.Round(
                    Proforma.TotalAmount
                    - Proforma.DiscountAmount
                    + Proforma.InsuranceFreightAmount
                    + Proforma.InsuranceFreightAmountII, 4);

                SelectedCustomer = AllCustomers.FirstOrDefault(c => c.CustCode == p.CustCode);
                if (SelectedCustomer != null)
                {
                    Countries = await ProformaService.GetCountriesForCustomerAsync(p.CustCode);
                }
                
                await LoadCustomerCountryLookups();
                await LoadCompanyLookups();

                if (OtherInfoTabRef != null)
                {
                    await OtherInfoTabRef.RestorePortSelections(p.LoadPort, p.DischargePort);
                }

                UpdateAmountInWords();
            }
            else
            {
                ShowError($"Proforma {PInvoice} not found.");
                GoBack();
            }
        }

        // ── Cascade Handlers ───────────────────────────────────────────────

        private async Task OnCompanyChanged()
        {
            await LoadCompanyLookups();
            if (NewItemsTabRef != null)
                await NewItemsTabRef.LoadAvailableOrders(Proforma.CompanyRefID, Proforma.CustCode, Proforma.Country);
        }

        private async Task LoadCompanyLookups()
        {
            if (Proforma.CompanyRefID > 0)
                CompanyBankAccounts = await ProformaService.GetCompanyBankAccountsAsync(Proforma.CompanyRefID);
            else
                CompanyBankAccounts.Clear();
        }

        private async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            SelectedCustomer = cust;
            Proforma.CustCode = cust?.CustCode ?? string.Empty;
            if (cust != null)
            {
                Countries = await ProformaService.GetCountriesForCustomerAsync(cust.CustCode);
                if (Countries.Count == 1)
                {
                    Proforma.Country = Countries[0];
                    await OnCountryChanged();
                }
                else
                {
                    Proforma.Country = string.Empty;
                }
            }
            else
            {
                Countries.Clear();
                Proforma.Country = string.Empty;
            }
        }

        private async Task OnCountryChanged()
        {
            if (!string.IsNullOrEmpty(Proforma.CustCode) && !string.IsNullOrEmpty(Proforma.Country))
            {
                // Auto-fill currency
                Proforma.Currency = await ProformaService.GetCurrencyForCustomerAsync(Proforma.CustCode, Proforma.Country);
                
                // Load defaults
                var defaults = await ProformaService.GetCustomerCountryDefaultsAsync(Proforma.CustCode, Proforma.Country);
                if (!string.IsNullOrEmpty(defaults.PaymentTerms))
                    Proforma.PaymentTerms = defaults.PaymentTerms;
                if (!string.IsNullOrEmpty(defaults.DefaultShipMethod))
                    Proforma.SMethod = defaults.DefaultShipMethod;

                await LoadCustomerCountryLookups();

                if (NewItemsTabRef != null)
                    await NewItemsTabRef.LoadAvailableOrders(Proforma.CompanyRefID, Proforma.CustCode, Proforma.Country);

                if (OtherInfoTabRef != null)
                    OtherInfoTabRef.SetDisCountryList(Countries); // Or could pass single country if desired
            }
        }

        private async Task LoadCustomerCountryLookups()
        {
            CustomerBanks = await ProformaService.GetCustomerBanksAsync(Proforma.CustCode, Proforma.Country);
            ShippingAddressPresets = await ProformaService.GetShippingAddressPresetsAsync(Proforma.CustCode, Proforma.Country);
            InvoiceToPresets = await ProformaService.GetInvoiceToPresetsAsync(Proforma.CustCode, Proforma.Country);
        }

        private void OnTradeTermsChanged()
        {
            if (Proforma.TradeTerms == "FOB-Karachi")
            {
                Proforma.SMethod = "By Sea";
                Proforma.InsuranceAgentID = null;
            }
            else if (Proforma.TradeTerms == "FOB-Sialkot" || Proforma.TradeTerms == "C&F-Sialkot" || Proforma.TradeTerms == "C&F-Lahore")
            {
                Proforma.SMethod = "By Air";
                Proforma.InsuranceAgentID = null;
            }
            // If CIF/CIP, Insurance agent is enabled (handled in razor via parameter)
        }

        // ── Tab Event Handlers ─────────────────────────────────────────────

        private void OnTotalsChanged()
        {
            UpdateAmountInWords();
        }

        private void HandleRemoveItem(ProformaLineItemModel item)
        {
            if (item.EntryID > 0)
                Proforma.DeletedLineItemIds.Add(item.EntryID);
        }

        private void HandleAddItems(List<OrderItemForProformaModel> newItems)
        {
            foreach (var ni in newItems)
            {
                LineItems.Add(new ProformaLineItemModel
                {
                    OrderItemID = ni.OrderItemID,
                    OrderNo = ni.OrderNo,
                    ItemCode = ni.ItemCode,
                    CompItemCode = ni.CompItemCode,
                    Description = ni.Description,
                    Unit = ni.Unit,
                    Price = ni.Price,
                    CustomPrice = ni.CustomPrice,
                    Qty = ni.QtyToAdd,
                    InvQty = ni.QtyToAdd, // Initial InvQty
                    DeliveryDTItem = ni.DeliveryDTItem,
                    QualityItem = ni.QualityItem,
                    StampsItem = ni.StampsItem
                });
            }
            
            DetailsTabRef?.RefreshTotals();
            ActiveTab = 1; // Switch back to details tab
            ShowSuccess($"Added {newItems.Count} items.");
        }

        private async Task HandleCalcSteelWeight(string pInvoiceStr)
        {
            if (!string.IsNullOrEmpty(pInvoiceStr) && IsEditMode)
            {
                try
                {
                    // Usually this queries VrptProformas which needs data saved first
                    Proforma.SteelWeight = await ProformaService.CalculateSteelWeightAsync(pInvoiceStr);
                    ShowSuccess("Steel weight calculated.");
                }
                catch (Exception ex)
                {
                    ShowError("Error calculating steel weight: " + ex.Message);
                }
            }
            else
            {
                ShowError("Please save the proforma first before calculating weight.");
            }
        }

        private void UpdateAmountInWords()
        {
            try
            {
                Proforma.AmountInWords = NumberToWordsConverter.Convert((decimal)Proforma.GrandTotal, Proforma.Currency);
            }
            catch
            {
                Proforma.AmountInWords = string.Empty;
            }
        }

        // ── Save & Navigation ──────────────────────────────────────────────

        private async Task SaveAndStay() => await SaveInternal(false);
        private async Task SaveAndNew()
        {
            if (await SaveInternal(true))
            {
                NavManager.NavigateTo("/export/proformas/new-proforma", forceLoad: true);
            }
        }
        private async Task SaveAndClose()
        {
            if (await SaveInternal(true))
            {
                GoBack();
            }
        }

        private async Task<bool> SaveInternal(bool navigateAway)
        {
            if (!Validate()) return false;

            IsLoading = true;
            try
            {
                Proforma.LineItems = LineItems;
                bool success = await ProformaService.SaveProformaInvoiceAsync(Proforma);
                
                if (success)
                {
                    ShowSuccess("Proforma saved successfully.");
                    if (!navigateAway && !IsEditMode)
                    {
                        // Switch to edit mode to stay
                        NavManager.NavigateTo($"/export/proformas/new-proforma/{Proforma.PInvoice}");
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                ShowError("Error saving: " + ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
            return false;
        }

        private bool Validate()
        {
            if (string.IsNullOrWhiteSpace(Proforma.CustCode)) { ShowError("Customer is required."); return false; }
            if (string.IsNullOrWhiteSpace(Proforma.Country)) { ShowError("Country is required."); return false; }
            if (string.IsNullOrWhiteSpace(Proforma.PInvoice)) { ShowError("Proforma # is required."); return false; }
            if (!LineItems.Any()) { ShowError("Cannot save an empty proforma. Add items first."); return false; }
            
            return true;
        }

        private async Task DeleteProforma()
        {
            if (!IsEditMode) return;
            
            // In a real app, use a proper confirmation dialog
            // Here assuming confirmed
            try
            {
                IsLoading = true;
                bool success = await ProformaService.DeleteProformaInvoiceAsync(Proforma.PInvoice);
                if (success)
                {
                    ShowSuccess("Proforma deleted.");
                    GoBack();
                }
            }
            catch (Exception ex)
            {
                ShowError("Error deleting: " + ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void GoBack()
        {
            NavManager.NavigateTo("/export/proformas/proforma-list");
        }

        private void ShowSuccess(string msg)
        {
            NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Success, Summary = "Success", Detail = msg, Duration = 3000 });
        }

        private void ShowError(string msg)
        {
            NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error", Detail = msg, Duration = 4000 });
        }
    }
}
