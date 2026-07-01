using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services.Export;
using Impulse.Services;

namespace Impulse.Pages.Export.Orders
{
    public partial class NewCustomerOrder : ComponentBase
    {
        [Parameter]
        public string? OrderNo { get; set; }

        [Inject]
        public ICustomerOrderService CustomerOrderService { get; set; } = default!;

        [Inject]
        public IDBHelperService DBHelperService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        public Radzen.NotificationService NotificationService { get; set; } = default!;

        // bound model
        public CustomerOrderViewModel Order { get; set; } = new();

        // list of deleted line items tracking
        public List<int> DeletedItemIds { get; set; } = new();

        // layout state
        public string ActiveTab { get; set; } = "ArticleDetail";
        public bool IsSaving { get; set; } = false;

        // lookup lists
        public List<CompanyLookupModel> Companies { get; set; } = new();
        public List<CustomerLookupModel> Customers { get; set; } = new();
        public List<string> Countries { get; set; } = new();
        public List<GenericDropDownModel> TradeTermsList { get; set; } = new();
        public List<GenericDropDownModel> PaymentTermsList { get; set; } = new();
        public List<GenericDropDownModel> ShippingMethods { get; set; } = new();
        public List<GenericDropDownModel> Qualities { get; set; } = new();
        public List<GenericDropDownModel> FinishingQualities { get; set; } = new();
        public List<CustomerOrderItemViewModel> CatalogItems { get; set; } = new();
        public List<string> QuotationsList { get; set; } = new();

        // Helpers for typeahead selectors
        public CompanyLookupModel? SelectedCompany { get; set; }
        public CustomerLookupModel? SelectedCustomer { get; set; }

        // Line-Item Add inputs
        public CustomerOrderItemViewModel NewItemInput { get; set; } = new();
        public CustomerOrderItemViewModel? SelectedCatalogItem { get; set; }
        public string SelectedQuotationNo { get; set; } = string.Empty;

        // Excel Import Lists
        public List<CustomerOrderItemViewModel> ExcelNICompany { get; set; } = new();
        public List<CustomerOrderItemViewModel> ExcelNICustomer { get; set; } = new();
        public List<CustomerOrderItemViewModel> ExcelOK { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();

            if (!string.IsNullOrEmpty(OrderNo))
            {
                var ord = await CustomerOrderService.GetOrderAsync(OrderNo);
                if (ord != null)
                {
                    Order = ord;
                    SelectedCompany = Companies.FirstOrDefault(c => c.EntryID == Order.CompanyRefID);
                    SelectedCustomer = Customers.FirstOrDefault(c => c.CustCode == Order.CustCode);
                    if (SelectedCustomer != null)
                    {
                        Countries = await CustomerOrderService.GetCountriesForCustomerAsync(SelectedCustomer.CustCode);
                        await HandleCustomerCountryChanged();
                    }
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Error",
                        Detail = "Order not found.",
                        Duration = 4000
                    });
                }
            }
            else
            {
                // Default values
                Order.DT = DateTime.Today;
                Order.DeliveryDT = DateTime.Today.AddMonths(2);
                Order.OrderRevisionDT = DateTime.Today;
                Order.OrderRevisionNo = 0;
                await RefreshInternalRef();
            }
        }

        private async Task RefreshInternalRef()
        {
            try
            {
                int companyId = SelectedCompany?.EntryID ?? 1;
                Order.InternalRefNo = await CustomerOrderService.GetNextInternalRefNoAsync(companyId);
                Order.OrderNo = Order.InternalRefNo;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to generate Internal Ref: {ex.Message}");
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
                Qualities = await CustomerOrderService.GetQualitiesAsync();
                FinishingQualities = await CustomerOrderService.GetFinishingQualitiesAsync();

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

        public async Task<IEnumerable<CompanyLookupModel>> SearchCompanies(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Companies;
            return await Task.FromResult(Companies.Where(c => c.CompanyName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public async Task OnCompanySelected(CompanyLookupModel? company)
        {
            SelectedCompany = company;
            if (company != null)
            {
                Order.CompanyRefID = company.EntryID;
                await RefreshInternalRef();
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
                Order.CustCode = customer.CustCode;
                Countries = await CustomerOrderService.GetCountriesForCustomerAsync(customer.CustCode);
                if (Countries.Any())
                {
                    Order.Country = Countries.First();
                    await HandleCustomerCountryChanged();
                }
            }
            else
            {
                Order.CustCode = string.Empty;
                Countries.Clear();
                Order.Country = string.Empty;
                CatalogItems.Clear();
                QuotationsList.Clear();
            }
        }

        public async Task OnCountryChanged(string country)
        {
            Order.Country = country;
            await HandleCustomerCountryChanged();
        }

        public async Task OnCountryChangedArgs(ChangeEventArgs e)
        {
            await OnCountryChanged(e.Value?.ToString() ?? string.Empty);
        }

        private async Task HandleCustomerCountryChanged()
        {
            if (!string.IsNullOrEmpty(Order.CustCode) && !string.IsNullOrEmpty(Order.Country))
            {
                Order.Currency = await CustomerOrderService.GetCurrencyForCustomerAsync(Order.CustCode, Order.Country);

                var defaults = await CustomerOrderService.GetCustomerDefaultsAsync(Order.CustCode, Order.Country);
                if (defaults != null)
                {
                    Order.PaymentTerms = defaults.PaymentTerms;
                    Order.TradeTerms = defaults.TradeTerms;
                    Order.StampDT = defaults.Stamps;
                    Order.OrderRcvdVia = defaults.RcvdVia;
                }

                CatalogItems = await CustomerOrderService.GetCustomerCatalogItemsAsync(Order.CustCode, Order.Country);
                QuotationsList = await CustomerOrderService.GetQuotationNumbersAsync(Order.CustCode, Order.Country);
            }
        }

        // ------------------------------------------------------------------ //
        //  Grid / Line Item Actions
        // ------------------------------------------------------------------ //

        public async Task OnCatalogItemChanged(CustomerOrderItemViewModel? catItem)
        {
            SelectedCatalogItem = catItem;
            if (catItem != null)
            {
                NewItemInput.ItemCode = catItem.ItemCode;
                NewItemInput.CompItemCode = catItem.CompItemCode;
                NewItemInput.Description = catItem.Description;
                NewItemInput.Unit = catItem.Unit;
                NewItemInput.Price = catItem.Price;
                NewItemInput.Weight = catItem.Weight;
            }
            else
            {
                NewItemInput.ItemCode = string.Empty;
                NewItemInput.CompItemCode = string.Empty;
                NewItemInput.Description = string.Empty;
                NewItemInput.Price = 0;
                NewItemInput.Weight = 0;
            }
            await Task.CompletedTask;
        }

        public async Task OnFinishingTypeChanged(int index)
        {
            NewItemInput.Item_Finishing_Type = index;
            if (SelectedCatalogItem != null && index > 0)
            {
                string field = index == 1 ? "Temper_Rate" : "First_Inspection_Rate";
                decimal customRate = await CustomerOrderService.GetItemCustomRateAsync(Order.CustCode, Order.Country, SelectedCatalogItem.CompItemCode, field);
                NewItemInput.Price = customRate;
            }
            else if (SelectedCatalogItem != null)
            {
                NewItemInput.Price = SelectedCatalogItem.Price;
            }
        }

        public void AddLineItem()
        {
            if (string.IsNullOrEmpty(NewItemInput.CompItemCode))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Select an item to add.",
                    Duration = 4000
                });
                return;
            }

            if (NewItemInput.Qty <= 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Quantity must be greater than zero.",
                    Duration = 4000
                });
                return;
            }

            var line = new CustomerOrderItemViewModel
            {
                ItemCode = NewItemInput.ItemCode,
                CompItemCode = NewItemInput.CompItemCode,
                Description = NewItemInput.Description,
                Price = NewItemInput.Price,
                Qty = NewItemInput.Qty,
                InvQty = NewItemInput.Qty,
                CustomPrice = NewItemInput.Price,
                Unit = NewItemInput.Unit,
                Quality = NewItemInput.Quality,
                Stamps = NewItemInput.Stamps,
                Remarks = NewItemInput.Remarks,
                Weight = NewItemInput.Weight,
                Item_Finishing_Type = NewItemInput.Item_Finishing_Type,
                Item_Finishing_Type_Text = NewItemInput.Item_Finishing_Type == 1 ? "Temper Inspection" : (NewItemInput.Item_Finishing_Type == 2 ? "First Inspection" : "Finished"),
                IW_OrderNo = NewItemInput.IW_OrderNo,
                IW_BatchNo = NewItemInput.IW_BatchNo,
                DeliveryDT = NewItemInput.DeliveryDT,
                SortNo = Order.OrderItems.Count + 1
            };

            Order.OrderItems.Add(line);

            // Reset inputs
            NewItemInput = new();
            SelectedCatalogItem = null;
        }

        public void RemoveLineItem(CustomerOrderItemViewModel item)
        {
            Order.OrderItems.Remove(item);
            if (item.ID > 0)
            {
                DeletedItemIds.Add(item.ID);
            }
            // Recalculate SortNo
            for (int i = 0; i < Order.OrderItems.Count; i++)
            {
                Order.OrderItems[i].SortNo = i + 1;
            }
        }

        public async Task LoadQuotation()
        {
            if (string.IsNullOrEmpty(SelectedQuotationNo)) return;
            try
            {
                var qItems = await CustomerOrderService.GetQuotationItemsAsync(SelectedQuotationNo);
                foreach (var qi in qItems)
                {
                    qi.SortNo = Order.OrderItems.Count + 1;
                    Order.OrderItems.Add(qi);
                }
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Quotation Loaded",
                    Detail = $"{qItems.Count} items appended to order.",
                    Duration = 4000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load quotation: {ex.Message}",
                    Duration = 4000
                });
            }
        }

        // ------------------------------------------------------------------ //
        //  Excel Import Actions
        // ------------------------------------------------------------------ //

        public async Task ImportExcelMock()
        {
            if (string.IsNullOrEmpty(Order.CustCode))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Select a Customer first.",
                    Duration = 4000
                });
                return;
            }

            try
            {
                await CustomerOrderService.ClearExcelImportTableAsync();

                // Mock Excel Row Insertions representing standard items
                await CustomerOrderService.InsertExcelImportRowAsync("02.790.00", "Periotomes Straight", 50, "Pcs", 3.25m, 0.035m, "SS", "UDMS1");
                await CustomerOrderService.InsertExcelImportRowAsync("99.999.99", "Unknown Item Missing Company Catalog", 20, "Pcs", 1.50m, 0.020m, "SS", "UDMS1");

                await CustomerOrderService.ProcessExcelImportAsync(Order.CustCode, Order.Country);

                ExcelNICompany = await CustomerOrderService.GetExcelImportErrorsAsync(true);
                ExcelNICustomer = await CustomerOrderService.GetExcelImportErrorsAsync(false);
                ExcelOK = await CustomerOrderService.GetExcelImportOkAsync(Order.CustCode, Order.Country);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Info,
                    Summary = "Excel Processed",
                    Detail = "Excel sheet loaded with validations.",
                    Duration = 4000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to process excel: {ex.Message}",
                    Duration = 4000
                });
            }
        }

        public void AppendExcelOkItems()
        {
            if (!ExcelOK.Any()) return;
            foreach (var item in ExcelOK)
            {
                item.SortNo = Order.OrderItems.Count + 1;
                Order.OrderItems.Add(item);
            }
            ExcelOK.Clear();
            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Success,
                Summary = "Items Added",
                Detail = "Validated excel rows appended to line items.",
                Duration = 4000
            });
        }

        // ------------------------------------------------------------------ //
        //  Save Operations
        // ------------------------------------------------------------------ //

        public async Task SaveOrderAndClose()
        {
            if (string.IsNullOrEmpty(Order.OrderNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Order No is required.",
                    Duration = 4000
                });
                return;
            }

            if (!Order.OrderItems.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "At least one item line must be added.",
                    Duration = 4000
                });
                return;
            }

            IsSaving = true;
            try
            {
                bool success = await CustomerOrderService.SaveOrderAsync(Order, DeletedItemIds);
                if (success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Saved",
                        Detail = "Customer Order saved successfully.",
                        Duration = 4000
                    });
                    Cancel();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Save failed: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public async Task SaveOrderAndStay()
        {
            if (string.IsNullOrEmpty(Order.OrderNo)) return;
            IsSaving = true;
            try
            {
                bool success = await CustomerOrderService.SaveOrderAsync(Order, DeletedItemIds);
                if (success)
                {
                    DeletedItemIds.Clear();
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Saved",
                        Detail = "Order saved.",
                        Duration = 3000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Save failed: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public async Task SaveOrderAndNew()
        {
            if (string.IsNullOrEmpty(Order.OrderNo)) return;
            IsSaving = true;
            try
            {
                bool success = await CustomerOrderService.SaveOrderAsync(Order, DeletedItemIds);
                if (success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Saved",
                        Detail = "Order saved. Clearing form.",
                        Duration = 3000
                    });
                    Order = new();
                    DeletedItemIds.Clear();
                    SelectedCustomer = null;
                    await RefreshInternalRef();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Save failed: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public void Cancel()
        {
            NavigationManager.NavigateTo("/export");
        }

        public void ChangeTab(string tabName)
        {
            ActiveTab = tabName;
        }

        public void AddRevision()
        {
            Order.OrderRevisionNo++;
            Order.OrderRevisionDT = DateTime.Today;
        }
    }
}
