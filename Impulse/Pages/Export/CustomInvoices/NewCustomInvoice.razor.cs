using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Components.Export.NewCustomInvoiceTabs;
using Impulse.Services.Export;
using Impulse.Utilities;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Export.CustomInvoices
{
    public partial class NewCustomInvoice
    {
        [Parameter] public string? CustomInvoiceNo { get; set; }

        [Inject] private ICustomInvoiceService InvoiceService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

        // ── Main Model ─────────────────────────────────────────────────────
        public CustomInvoiceViewModel Invoice { get; set; } = new();
        
        // ── State ──────────────────────────────────────────────────────────
        private bool IsEditMode => !string.IsNullOrEmpty(CustomInvoiceNo);
        private bool IsLoading = false;
        private int ActiveTab = 0; // Starts at 0: Invoice Items

        // ── Tab refs ──────────────────────────────────────────────────────
        private OtherInfoTab? OtherInfoTabRef;

        // ── Lookups ────────────────────────────────────────────────────────
        private List<GenericDropDownModel> Companies = new();
        private List<CustomerLookupModel> AllCustomers = new();
        private CustomerLookupModel? SelectedCustomer;
        private List<string> Countries = new();
        private List<GenericDropDownModel> TradeTermsList = new();
        private List<GenericDropDownModel> PaymentTermsList = new();
        private List<GenericDropDownModel> ShippingMethods = new();
        
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
                    await LoadCustomInvoice();
                }
                else
                {
                    Invoice.CompanyRefID = int.TryParse(Companies.FirstOrDefault()?.DropDownValue_ID, out int compId) ? compId : 1;
                    // Legacy logic for "IAA"
                    var comp = Companies.FirstOrDefault(c => c.DropDownValue_ID == Invoice.CompanyRefID.ToString())?.DropDownValue_Description ?? "IAA";
                    Invoice.CustomInvoiceNo = await InvoiceService.GetNextCustomInvoiceNumberAsync(comp);
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
            Companies = await InvoiceService.GetCompaniesAsync();
            AllCustomers = await InvoiceService.GetCustomersAsync();
            TradeTermsList = await InvoiceService.GetTradeTermsAsync();
            PaymentTermsList = await InvoiceService.GetPaymentTermsAsync();
            ShippingMethods = await InvoiceService.GetShippingMethodsAsync();
            ShippingAgents = await InvoiceService.GetShippingAgentsAsync();
            InsuranceAgents = await InvoiceService.GetInsuranceAgentsAsync();
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            await Task.Delay(1); // Yield
            if (string.IsNullOrWhiteSpace(searchText))
                return AllCustomers;
            return AllCustomers.Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task LoadCustomInvoice()
        {
            var data = await InvoiceService.GetCustomInvoiceAsync(CustomInvoiceNo!);
            if (data == null)
            {
                ShowError("Invoice not found.");
                GoBack();
                return;
            }
            Invoice = data;
            
            SelectedCustomer = AllCustomers.FirstOrDefault(c => c.CustCode == Invoice.CustCode);
            if (SelectedCustomer != null)
            {
                Countries = await InvoiceService.GetCountriesForCustomerAsync(SelectedCustomer.CustCode);
            }
            await LoadDependentLookups();
            CalculateTotals();

            // Restore port selections (must be after rendering)
            await Task.Yield(); // allow Blazor to render
            if (OtherInfoTabRef != null)
            {
                OtherInfoTabRef.SetDisChargeCountryList(Countries);
                await OtherInfoTabRef.RestorePortSelectionsAsync();
            }
        }

        private async Task OnCompanyChanged()
        {
            CompanyBankAccounts = await InvoiceService.GetCompanyBankAccountsAsync(Invoice.CompanyRefID);
            
            // Recalculate invoice number if new mode
            if (!IsEditMode)
            {
                var comp = Companies.FirstOrDefault(c => c.DropDownValue_ID == Invoice.CompanyRefID.ToString())?.DropDownValue_Description ?? "IAA";
                Invoice.CustomInvoiceNo = await InvoiceService.GetNextCustomInvoiceNumberAsync(comp);
            }
        }

        private async Task OnCustomerChanged(CustomerLookupModel? customer)
        {
            SelectedCustomer = customer;
            if (SelectedCustomer == null)
            {
                Invoice.CustCode = string.Empty;
                Invoice.Country = string.Empty;
                Countries.Clear();
                return;
            }

            Invoice.CustCode = SelectedCustomer.CustCode;
            Countries = await InvoiceService.GetCountriesForCustomerAsync(Invoice.CustCode);
            OtherInfoTabRef?.SetDisChargeCountryList(Countries);
            
            if (Countries.Count == 1)
            {
                Invoice.Country = Countries.First();
                await OnCountryChanged();
            }
        }

        private async Task OnCountryChanged()
        {
            if (string.IsNullOrEmpty(Invoice.CustCode) || string.IsNullOrEmpty(Invoice.Country)) return;
            
            Invoice.Currency = await InvoiceService.GetCurrencyForCustomerAsync(Invoice.CustCode, Invoice.Country);
            await LoadDependentLookups();
            
            // Update the discharge country list in OtherInfoTab
            OtherInfoTabRef?.SetDisChargeCountryList(Countries);
            
            if (!IsEditMode)
            {
                // Auto-fill addresses if we have exactly 1
                if (ShippingAddressPresets.Count == 1) Invoice.ComAgainst = ShippingAddressPresets[0].ShippingAddress;
                if (InvoiceToPresets.Count == 1) Invoice.InvoiceTo = InvoiceToPresets[0].InvoiceTo;
                
                // Fetch advance payments
                Invoice.AdvancePayments = await InvoiceService.GetAdvancePaymentsAsync(Invoice.CustCode);
            }
        }


        private async Task LoadDependentLookups()
        {
            CustomerBanks = await InvoiceService.GetCustomerBanksAsync(Invoice.CustCode, Invoice.Country);
            ShippingAddressPresets = await InvoiceService.GetShippingAddressPresetsAsync(Invoice.CustCode, Invoice.Country);
            InvoiceToPresets = await InvoiceService.GetInvoiceToPresetsAsync(Invoice.CustCode, Invoice.Country);
            await OnCompanyChanged(); // Ensure company banks are loaded
        }

        private void OnTotalsChanged()
        {
            CalculateTotals();
            StateHasChanged();
        }

        private async Task HandleAddItems((List<CustomPListModel> Items, string ProformaNo) data)
        {
            var selectedItems = data.Items;
            var proformaNo = data.ProformaNo;

            if (Invoice.LineItems.Count == 0 && !string.IsNullOrEmpty(proformaNo))
            {
                // Auto-populate headers from Proforma if this is the first items being added
                var proformaHeader = await InvoiceService.GetProformaHeaderAsync(proformaNo);
                if (proformaHeader != null)
                {
                    Invoice.TradeTerms = proformaHeader.TradeTerms ?? Invoice.TradeTerms;
                    Invoice.PartialShipment = proformaHeader.PartialShipment ?? Invoice.PartialShipment;
                    Invoice.TransShipment = proformaHeader.TransShipment ?? Invoice.TransShipment;
                    Invoice.PaymentTerms = proformaHeader.PaymentTerms ?? Invoice.PaymentTerms;
                    Invoice.Packaging = proformaHeader.Packaging ?? Invoice.Packaging;
                    Invoice.ShippingMode = proformaHeader.ShippingMode ?? Invoice.ShippingMode;
                    Invoice.CustomDT = proformaHeader.CustomDT ?? Invoice.CustomDT;
                    Invoice.InvoiceOf = proformaHeader.InvoiceOf ?? Invoice.InvoiceOf;
                    Invoice.ShippingAgentID = proformaHeader.ShippingAgentID ?? Invoice.ShippingAgentID;
                    Invoice.OurBankID = proformaHeader.OurBankID ?? Invoice.OurBankID;
                    Invoice.LoadPortID = proformaHeader.LoadPortID ?? Invoice.LoadPortID;
                    Invoice.DisPortID = proformaHeader.DisPortID ?? Invoice.DisPortID;
                    Invoice.Consignee = proformaHeader.Consignee ?? Invoice.Consignee;
                    Invoice.ComNotify = proformaHeader.ComNotify ?? Invoice.ComNotify;
                    Invoice.ComConsignee = proformaHeader.ComConsignee ?? Invoice.ComConsignee;
                    Invoice.ComPayTerms = proformaHeader.ComPayTerms ?? Invoice.ComPayTerms;
                }

                var otherCharges = await InvoiceService.GetProformaOtherChargesAsync(proformaNo);
                if (otherCharges != null && otherCharges.Any())
                {
                    // Add other charges if empty
                    if (Invoice.OtherCharges.Count == 0)
                    {
                        foreach (var charge in otherCharges)
                        {
                            Invoice.OtherCharges.Add(new CustomInvoiceOtherChargeModel
                            {
                                SNo = charge.SNo,
                                ChargesDetail = charge.ChargesDetail,
                                ChargesAmt = charge.ChargesAmt,
                                Custom = false
                            });
                        }
                    }
                }
            }

            foreach (var item in selectedItems)
            {
                int totalQtyForThisEntry = ((item.CartonTo - item.CartonFrom) + 1) * item.Qty;
                
                // Add to InvoiceItems
                var existingItem = Invoice.LineItems.FirstOrDefault(x => x.RefID == item.OrderItemID);
                if (existingItem != null)
                {
                    existingItem.Qty += totalQtyForThisEntry;
                }
                else
                {
                    Invoice.LineItems.Add(new CustomInvoiceItemModel
                    {
                        RefID = item.OrderItemID,
                        ProformaInvoice = item.ProformaInvoice,
                        OrderNo = item.OrderNo,
                        ItemCode = item.ItemCode,
                        Description = item.Description,
                        Qty = totalQtyForThisEntry,
                        Price = 0, // Legacy logic often resets price or keeps it
                        CustomPrice = 0
                    });
                }

                // Add to Packing List
                item.ID = 0; // mark as new
                Invoice.PackingListItems.Add(item);
            }
            CalculateTotals();
            // Removed ActiveTab = 1 so the user stays on the Proforma Items tab to continue entering data
        }

        private void CalculateTotals()
        {
            Invoice.TotalAmount = Invoice.LineItems.Sum(x => x.Amount);
            Invoice.GrandTotal = Invoice.TotalAmount + Invoice.OtherCharges.Sum(x => x.ChargesAmt);
            Invoice.AmountInWords = NumberToWordsConverter.Convert((decimal)Invoice.GrandTotal, Invoice.Currency);
        }

        private async Task SaveAndStay() => await SaveInvoice(false);
        private async Task SaveAndClose() => await SaveInvoice(true);
        
        private async Task SaveAndNew()
        {
            if (await SaveInvoice(false))
            {
                NavManager.NavigateTo("/export/custominvoices/new-custominvoice", forceLoad: true);
            }
        }

        private async Task<bool> SaveInvoice(bool closeAfterSave)
        {
            if (string.IsNullOrWhiteSpace(Invoice.CustomInvoiceNo))
            {
                ShowError("Invoice No is required.");
                return false;
            }
            if (string.IsNullOrWhiteSpace(Invoice.CustCode))
            {
                ShowError("Customer is required.");
                return false;
            }
            if (!Invoice.LineItems.Any())
            {
                ShowError("At least one item is required.");
                return false;
            }

            IsLoading = true;
            try
            {
                Invoice.DeletedLineItemIds = Invoice.LineItems.Where(x => x.Qty <= 0 && x.EntryID > 0).Select(x => x.EntryID).ToList();
                Invoice.LineItems.RemoveAll(x => x.Qty <= 0); // Cleanup

                bool success = await InvoiceService.SaveCustomInvoiceAsync(Invoice);
                if (success)
                {
                    ShowSuccess("Saved successfully.");
                    if (closeAfterSave)
                        GoBack();
                    else if (!IsEditMode)
                        NavManager.NavigateTo($"/export/custominvoices/new-custominvoice/{Invoice.CustomInvoiceNo}");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                ShowError("Failed to save: " + ex.Message);
                return false;
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task DeleteInvoice()
        {
            if (!IsEditMode) return;
            IsLoading = true;
            try
            {
                if (await InvoiceService.DeleteCustomInvoiceAsync(Invoice.CustomInvoiceNo))
                {
                    ShowSuccess("Deleted successfully.");
                    GoBack();
                }
            }
            catch (Exception ex)
            {
                ShowError("Failed to delete: " + ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void GoBack()
        {
            // NavManager.NavigateTo("/export/custominvoices"); // Adjust route
            NavManager.NavigateTo("/"); // fallback for now
        }

        private void ShowSuccess(string msg) => NotificationService.Notify(NotificationSeverity.Success, "Success", msg, 3000);
        private void ShowError(string msg) => NotificationService.Notify(NotificationSeverity.Error, "Error", msg, 6000);
    }
}
