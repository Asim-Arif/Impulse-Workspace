using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components.Authorization;
using System.Security.Claims;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.JSInterop;
using BlazorContextMenu;

namespace Impulse.Pages.Export.Orders
{
    public partial class ArticlewiseShippedStatus : ComponentBase
    {
        [Inject] private IArticlewiseShippedStatusService StatusService { get; set; } = null!;
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = null!;
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IConfiguration Config { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;

        // Filters
        private List<CustomerLookupModel> customers = new();
        private CustomerLookupModel? selectedCustomer = null;

        private List<string> countries = new();
        private string selectedCountry = string.Empty;

        private List<ItemLookupModel> articles = new();
        private ItemLookupModel? selectedArticle = null;

        // Data grids
        private IEnumerable<ArticleShippedInvoiceModel> shippedInvoices = new List<ArticleShippedInvoiceModel>();
        private IEnumerable<ArticleUnshippedOrderModel> unshippedOrders = new List<ArticleUnshippedOrderModel>();
        private IEnumerable<ArticleCustomerPriceModel> customerPrices = new List<ArticleCustomerPriceModel>();
        private IEnumerable<ArticleQuotationModel> quotations = new List<ArticleQuotationModel>();

        // Password Authorization
        private bool showPasswordModal = false;
        private string enteredPassword = string.Empty;
        private string passwordError = string.Empty;
        private Func<Task>? pendingSaveAction = null;
        private string currentUserName = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            currentUserName = authState.User.Identity?.Name ?? "Administrator";

            customers = await CustomerOrderService.GetCustomersAsync();
            
            // Insert dummy "All" customer if needed, but BlazoredTypeahead allows null
            await LoadArticlesAsync();
        }

        public async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string search)
        {
            if (string.IsNullOrWhiteSpace(search)) return customers;
            return await Task.FromResult(customers.Where(c => c.CustCode.Contains(search, StringComparison.OrdinalIgnoreCase) || 
                                                            c.Name.Contains(search, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public async Task<IEnumerable<ItemLookupModel>> SearchArticles(string search)
        {
            if (string.IsNullOrWhiteSpace(search)) return articles.Take(100);
            return await Task.FromResult(articles.Where(a => a.ItemName.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                                                             a.ItemID.Contains(search, StringComparison.OrdinalIgnoreCase)).Take(100).ToList());
        }

        private async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            selectedCustomer = cust;
            
            if (cust != null)
            {
                // Fetch distinct countries for this customer from ForeignCustomers
                using var conn = new SqlConnection(Config.GetConnectionString("DefaultConnection"));
                var custCountries = await conn.QueryAsync<string>("SELECT DISTINCT Country FROM ForeignCustomers WHERE CustCode = @CustCode", new { CustCode = cust.CustCode });
                countries = custCountries.Where(c => !string.IsNullOrEmpty(c)).ToList();
                selectedCountry = countries.FirstOrDefault() ?? string.Empty;
            }
            else
            {
                countries.Clear();
                selectedCountry = string.Empty;
            }

            await LoadArticlesAsync();
            ClearData();
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? string.Empty;
            await LoadArticlesAsync();
            ClearData();
        }

        private async Task OnArticleChanged(ItemLookupModel? article)
        {
            selectedArticle = article;
            if (selectedArticle != null)
            {
                await LoadData();
            }
            else
            {
                ClearData();
            }
        }

        private async Task LoadArticlesAsync()
        {
            using var conn = new SqlConnection(Config.GetConnectionString("DefaultConnection"));
            if (selectedCustomer == null || selectedCustomer.CustCode == "0")
            {
                // Load from Items
                string sql = "SELECT ItemID, ItemID as CompItemCode, ItemName FROM Items WHERE InActive = 0 ORDER BY ItemID";
                articles = (await conn.QueryAsync<ItemLookupModel>(sql)).ToList();
            }
            else
            {
                // Load from VFCustomerCatalog based on CustCode and Country
                string sql = @"SELECT CompItemID as ItemID, CompItemID as CompItemCode, 
                                     '{' + CompItemID + '} ' + ItemID + ' {' + ISNULL(OCR,'') + '} {' + ISNULL(OCR2,'') + '} ' + CompDesc as ItemName 
                               FROM VFCustomerCatalog 
                               WHERE CustCode = @CustCode AND Country = @Country";
                articles = (await conn.QueryAsync<ItemLookupModel>(sql, new { CustCode = selectedCustomer.CustCode, Country = selectedCountry })).ToList();
            }
            selectedArticle = null;
        }

        private void ClearData()
        {
            shippedInvoices = new List<ArticleShippedInvoiceModel>();
            unshippedOrders = new List<ArticleUnshippedOrderModel>();
            customerPrices = new List<ArticleCustomerPriceModel>();
            quotations = new List<ArticleQuotationModel>();
        }

        private async Task LoadData()
        {
            if (selectedArticle == null) return;

            try
            {
                string itemCode = selectedArticle.ItemID; // For Items, ItemID is the CompItemCode. For VFCustomerCatalog, we selected CompItemID as ItemID
                string custCode = selectedCustomer?.CustCode ?? "0";

                shippedInvoices = await StatusService.GetShippedInvoicesAsync(itemCode, custCode);
                unshippedOrders = await StatusService.GetUnshippedOrdersAsync(itemCode, custCode);
                customerPrices = await StatusService.GetCustomerPricesAsync(itemCode);
                quotations = await StatusService.GetQuotationsAsync(itemCode);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error loading data", ex.Message);
            }
        }

        // Inline Editing Logic
        private void BeginEditPrice(ArticleCustomerPriceModel cp, string columnName, double value)
        {
            foreach (var item in customerPrices)
            {
                item.IsEditing = false;
            }

            cp.IsEditing = true;
            cp.EditingColumn = columnName;
            cp.EditPrice = value;
        }

        private void CancelEdit()
        {
            foreach (var item in customerPrices)
            {
                item.IsEditing = false;
            }
        }

        private async Task HandlePriceKeyUp(KeyboardEventArgs e, ArticleCustomerPriceModel cp)
        {
            if (e.Key == "Enter")
            {
                // Prompt for password
                enteredPassword = string.Empty;
                passwordError = string.Empty;
                showPasswordModal = true;

                pendingSaveAction = async () =>
                {
                    try
                    {
                        await StatusService.UpdateCustomerPriceAsync(cp.EntryID, cp.EditingColumn, cp.EditPrice, currentUserName);
                        
                        // Update UI Model directly to avoid full reload
                        var prop = cp.GetType().GetProperty(cp.EditingColumn);
                        if (prop != null && prop.CanWrite)
                        {
                            prop.SetValue(cp, cp.EditPrice);
                        }

                        NotificationServiceManager.ShowSuccess("Success", "Price updated successfully.");
                        CancelEdit();
                    }
                    catch (Exception ex)
                    {
                        NotificationServiceManager.ShowError("Update Failed", ex.Message);
                    }
                };
            }
            else if (e.Key == "Escape")
            {
                CancelEdit();
            }
        }

        private async Task VerifyPasswordAndSave()
        {
            passwordError = string.Empty;
            if (string.IsNullOrWhiteSpace(enteredPassword))
            {
                passwordError = "Password cannot be empty.";
                return;
            }

            string expectedPassword = await DbHelper.getDatabasePasswordAsync("EditPrice"); // Reusing standard setting
            if (enteredPassword != expectedPassword)
            {
                passwordError = "Invalid password.";
                return;
            }

            showPasswordModal = false;

            if (pendingSaveAction != null)
            {
                await pendingSaveAction.Invoke();
                pendingSaveAction = null;
            }
        }

        // Context Menu & Reports Logic
        private async Task OpenOptionsBarMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsMenu", (int)e.ClientX, (int)e.ClientY);
        }

        private async Task PrintShipped(ItemClickEventArgs e)
        {
            if (selectedArticle != null)
            {
                var request = new Impulse.Services.ReportRequest
                {
                    ReportName = "rptArticlewiseSales.rpt",
                    SelectionFormula = $"{{VArticlewiseSales.ItemID}}='{selectedArticle.ItemID}'"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
        }

        private async Task PrintUnShipped(ItemClickEventArgs e)
        {
            if (selectedArticle != null)
            {
                var request = new Impulse.Services.ReportRequest
                {
                    ReportName = "rptPartiallyShippedOrders.rpt",
                    SelectionFormula = $"{{Items.ItemID}}='{selectedArticle.ItemID}' AND NOT ({{FCustomerOrders.OrderNo}} IN['Stock-Order','Stock-OrderIss'])"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
        }
    }
}
