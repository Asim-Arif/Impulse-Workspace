using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;
using Impulse.Services.Export;
using Impulse.Services.Company;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Accounts;

namespace Impulse.Pages.Export.Quotations
{
    public partial class NewCustomerQuotation : ComponentBase
    {
        [Inject] private ICustomerQuotationService CustomerQuotationService { get; set; } = default!;
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private IBankListDataAccess BankListDataAccess { get; set; } = default!;

        [Parameter] public int? QuotationNo { get; set; }

        public CustomerQuotationViewModel Quotation { get; set; } = new();
        public bool IsSaving { get; set; } = false;
        public string ActiveTab { get; set; } = "ArticleDetail";

        // Lookups
        public List<CompanyLookupModel> Companies { get; set; } = new();
        public List<CustomerLookupModel> Customers { get; set; } = new();
        public List<string> Countries { get; set; } = new();
        public List<GenericDropDownModel> TradeTermsList { get; set; } = new();
        public List<GenericDropDownModel> PaymentTermsList { get; set; } = new();
        public List<GenericDropDownModel> ShippingMethods { get; set; } = new();
        public List<GenericDropDownModel> FinishingQualities { get; set; } = new();
        public List<BankListModel> Banks { get; set; } = new();
        public List<CustomerOrderItemViewModel> CatalogItems { get; set; } = new();

        public CompanyLookupModel? SelectedCompany { get; set; }
        public CustomerLookupModel? SelectedCustomer { get; set; }
        
        // Items State
        public CustomerQuotationItemViewModel NewItemInput { get; set; } = new();
        public CustomerQuotationItemViewModel? EditingItem { get; set; }
        public CustomerOrderItemViewModel? SelectedCatalogItem { get; set; }
        public List<int> DeletedItemIds { get; set; } = new();

        // Other Charges State
        public CustomerQuotationOtherChargeViewModel NewChargeInput { get; set; } = new();
        public CustomerQuotationOtherChargeViewModel? EditingCharge { get; set; }
        public List<int> DeletedChargeIds { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();

            if (QuotationNo.HasValue && QuotationNo.Value > 0)
            {
                await LoadQuotation(QuotationNo.Value);
            }
            else
            {
                await RefreshQuotationNo();
            }
            RefreshChargeSNo();
        }

        private async Task RefreshQuotationNo()
        {
            try
            {
                Quotation.QuotationNo = await CustomerQuotationService.GetNextQuotationNoAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to generate Quotation No: {ex.Message}");
            }
        }

        private async Task LoadLookups()
        {
            try
            {
                var rawCompanies = await CustomerOrderService.GetCompaniesAsync();
                Companies = rawCompanies.Select(c => new CompanyLookupModel 
                { 
                    EntryID = int.TryParse(c.DropDownValue_ID, out int val) ? val : 1, 
                    CompanyName = c.DropDownValue_Description 
                }).ToList();

                Customers = await CustomerOrderService.GetCustomersAsync();
                TradeTermsList = await CustomerOrderService.GetTradeTermsAsync();
                PaymentTermsList = await CustomerOrderService.GetPaymentTermsAsync();
                ShippingMethods = await CustomerOrderService.GetShippingMethodsAsync();
                FinishingQualities = await CustomerOrderService.GetFinishingQualitiesAsync();
                Banks = await BankListDataAccess.GetBankList();

                if (Companies.Any()) SelectedCompany = Companies.First();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load lookup configurations: {ex.Message}",
                    Duration = 4000
                });
            }
        }

        private async Task LoadQuotation(int quotationNo)
        {
            try
            {
                var data = await CustomerQuotationService.GetQuotationAsync(quotationNo);
                if (data != null)
                {
                    Quotation = data;

                    SelectedCompany = Companies.FirstOrDefault(c => c.EntryID == Quotation.CompanyRefID);
                    SelectedCustomer = Customers.FirstOrDefault(c => c.CustCode == Quotation.CustCode);
                    
                    if (SelectedCustomer != null)
                    {
                        Countries = await CustomerOrderService.GetCountriesForCustomerAsync(Quotation.CustCode);
                        await HandleCustomerCountryChanged(true);
                    }
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Warning, Summary = "Not Found", Detail = "Quotation not found" });
                    NavigationManager.NavigateTo("/export/customer-quotation-list");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Error, Summary = "Error", Detail = ex.Message });
            }
        }

        public async Task<IEnumerable<CompanyLookupModel>> SearchCompanies(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Companies;
            return await Task.FromResult(Companies.Where(c => c.CompanyName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public void OnCompanySelected(CompanyLookupModel? company)
        {
            SelectedCompany = company;
            if (company != null)
            {
                Quotation.CompanyRefID = company.EntryID;
            }
        }

        public async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Customers;
            return await Task.FromResult(Customers.Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public async Task OnCustomerSelected(CustomerLookupModel? customer)
        {
            SelectedCustomer = customer;
            if (customer != null)
            {
                Quotation.CustCode = customer.CustCode;
                Countries = await CustomerOrderService.GetCountriesForCustomerAsync(customer.CustCode);
                if (Countries.Any())
                {
                    Quotation.Country = Countries.First();
                    await HandleCustomerCountryChanged();
                }
            }
            else
            {
                Quotation.CustCode = string.Empty;
                Countries.Clear();
                Quotation.Country = string.Empty;
                CatalogItems.Clear();
            }
        }

        public async Task OnCountryChanged(string country)
        {
            Quotation.Country = country;
            await HandleCustomerCountryChanged();
        }

        public async Task OnCountryChangedArgs(ChangeEventArgs e)
        {
            await OnCountryChanged(e.Value?.ToString() ?? string.Empty);
        }

        private async Task HandleCustomerCountryChanged(bool isEdit = false)
        {
            if (!string.IsNullOrEmpty(Quotation.CustCode) && !string.IsNullOrEmpty(Quotation.Country))
            {
                Quotation.Currency = await CustomerOrderService.GetCurrencyForCustomerAsync(Quotation.CustCode, Quotation.Country);

                var defaults = await CustomerOrderService.GetCustomerDefaultsAsync(Quotation.CustCode, Quotation.Country);
                if (defaults != null && !isEdit)
                {
                    Quotation.PaymentTerms = defaults.PaymentTerms;
                    Quotation.TradeTerms = defaults.TradeTerms;
                }

                CatalogItems = await CustomerOrderService.GetCustomerCatalogItemsAsync(Quotation.CustCode, Quotation.Country);
            }
        }

        public void ChangeTab(string tabName)
        {
            ActiveTab = tabName;
        }

        // ------------------------------------------------------------------ //
        //  Items Actions
        // ------------------------------------------------------------------ //

        public async Task OnCatalogItemChanged(CustomerOrderItemViewModel? catItem)
        {
            SelectedCatalogItem = catItem;
            if (catItem != null)
            {
                NewItemInput.ItemCode = catItem.ItemCode;
                NewItemInput.ItemName = catItem.Description;
                NewItemInput.Price = catItem.Price;
            }
            else
            {
                NewItemInput.ItemCode = string.Empty;
                NewItemInput.ItemName = string.Empty;
                NewItemInput.Price = 0;
            }
            await Task.CompletedTask;
        }

        public void OnFinishingTypeChanged(int index)
        {
            if (index >= 0 && index < FinishingQualities.Count)
            {
                NewItemInput.QualityType = FinishingQualities[index].DropDownValue_Description;
            }
        }

        public void SelectForEdit(CustomerQuotationItemViewModel item)
        {
            EditingItem = item;
            NewItemInput = new CustomerQuotationItemViewModel
            {
                EntryID = item.EntryID,
                QuotationNo = item.QuotationNo,
                ItemCode = item.ItemCode,
                ItemName = item.ItemName,
                Price = item.Price,
                Qty = item.Qty,
                QualityType = item.QualityType,
                DeliveryDT = item.DeliveryDT,
                ItemRemarks = item.ItemRemarks
            };
            SelectedCatalogItem = CatalogItems.FirstOrDefault(c => c.ItemCode == item.ItemCode);
        }

        public void CancelEdit()
        {
            EditingItem = null;
            NewItemInput = new();
            SelectedCatalogItem = null;
        }

        public void AddLineItem()
        {
            if (string.IsNullOrWhiteSpace(NewItemInput.ItemCode)) return;

            if (EditingItem != null)
            {
                EditingItem.ItemCode = NewItemInput.ItemCode;
                EditingItem.ItemName = NewItemInput.ItemName;
                EditingItem.Price = NewItemInput.Price;
                EditingItem.Qty = NewItemInput.Qty;
                EditingItem.QualityType = NewItemInput.QualityType;
                EditingItem.DeliveryDT = NewItemInput.DeliveryDT;
                EditingItem.ItemRemarks = NewItemInput.ItemRemarks;
                EditingItem = null;
            }
            else
            {
                Quotation.QuotationItems.Add(new CustomerQuotationItemViewModel
                {
                    ItemCode = NewItemInput.ItemCode,
                    ItemName = NewItemInput.ItemName,
                    Price = NewItemInput.Price,
                    Qty = NewItemInput.Qty,
                    QualityType = NewItemInput.QualityType,
                    DeliveryDT = NewItemInput.DeliveryDT,
                    ItemRemarks = NewItemInput.ItemRemarks
                });
            }
            NewItemInput = new();
            SelectedCatalogItem = null;
        }

        public void RemoveLineItem(CustomerQuotationItemViewModel item)
        {
            if (item.EntryID > 0) DeletedItemIds.Add(item.EntryID);
            Quotation.QuotationItems.Remove(item);
        }

        // ------------------------------------------------------------------ //
        //  Other Charges Actions
        // ------------------------------------------------------------------ //
        public void SelectChargeForEdit(CustomerQuotationOtherChargeViewModel charge)
        {
            EditingCharge = charge;

            NewChargeInput.SNo = charge.SNo;
            NewChargeInput.ChargesDetail = charge.ChargesDetail;
            NewChargeInput.ChargesAmt = charge.ChargesAmt;
        }

        public void CancelEditCharge()
        {
            EditingCharge = null;
            NewChargeInput = new();
            RefreshChargeSNo();
        }

        public void AddCharge()
        {
            if (string.IsNullOrWhiteSpace(NewChargeInput.ChargesDetail)) return;

            if (EditingCharge != null)
            {
                EditingCharge.SNo = NewChargeInput.SNo;
                EditingCharge.ChargesDetail = NewChargeInput.ChargesDetail;
                EditingCharge.ChargesAmt = NewChargeInput.ChargesAmt;
                EditingCharge = null;
            }
            else
            {
                Quotation.OtherCharges.Add(new CustomerQuotationOtherChargeViewModel
                {
                    SNo = NewChargeInput.SNo,
                    ChargesDetail = NewChargeInput.ChargesDetail,
                    ChargesAmt = NewChargeInput.ChargesAmt
                });
            }
            NewChargeInput = new();
            RefreshChargeSNo();
        }

        public void RemoveCharge(CustomerQuotationOtherChargeViewModel charge)
        {
            if (charge.EntryID > 0) DeletedChargeIds.Add(charge.EntryID);
            Quotation.OtherCharges.Remove(charge);
            RefreshChargeSNo();
        }

        private void RefreshChargeSNo()
        {
            NewChargeInput.SNo = Quotation.OtherCharges.Any() ? Quotation.OtherCharges.Max(c => c.SNo) + 1 : 1;
        }

        // ------------------------------------------------------------------ //
        //  Save / Cancel logic
        // ------------------------------------------------------------------ //

        public void Cancel()
        {
            NavigationManager.NavigateTo("/export/customer-quotation-list");
        }

        public async Task SaveQuotationAndStay()
        {
            await SaveInternal();
        }

        public async Task SaveQuotationAndNew()
        {
            if (await SaveInternal())
            {
                NavigationManager.NavigateTo("/export/new-customer-quotation", forceLoad: true);
            }
        }

        public async Task SaveQuotationAndClose()
        {
            if (await SaveInternal())
            {
                NavigationManager.NavigateTo("/export/customer-quotation-list");
            }
        }

        private async Task<bool> SaveInternal()
        {
            if (string.IsNullOrWhiteSpace(Quotation.CustCode))
            {
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Error, Summary = "Validation", Detail = "Please select a Customer." });
                return false;
            }

            IsSaving = true;
            try
            {
                bool saved = await CustomerQuotationService.SaveQuotationAsync(Quotation, DeletedItemIds, DeletedChargeIds);
                if (saved)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Saved",
                        Detail = "Quotation saved successfully.",
                        Duration = 4000
                    });
                    
                    DeletedItemIds.Clear();
                    DeletedChargeIds.Clear();
                    
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Error, Summary = "Error", Detail = ex.Message });
                return false;
            }
            finally
            {
                IsSaving = false;
            }
        }
    }
}
