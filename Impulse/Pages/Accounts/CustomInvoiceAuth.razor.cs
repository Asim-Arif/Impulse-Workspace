using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;

namespace Impulse.Pages.Accounts
{
    public partial class CustomInvoiceAuth : ComponentBase
    {
        [Inject] private ICustomInvoiceAuthService CustomInvoiceAuthService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = true;
        private bool isAuthorizing = false;

        // Filters (legacy: cmbCust / cmbCountry / cmbRange / DtFrom / DtTo)
        private CustomerLookupModel? selectedCustomer;
        private string selectedCountry = string.Empty;
        private string selectedDateRange = "180";   // Legacy Form_Load: cmbRange.ListIndex = 5 (Last 180 Days)
        private DateTime fromDate = DateTime.Today.AddDays(-180);
        private DateTime toDate = DateTime.Today;

        // Authorization bar (legacy: CmbAccTitle / DTVchr)
        private List<ChartOfAccountsModel> salesAccounts = new();
        private ChartOfAccountsModel? selectedSalesAccount;
        private DateTime voucherDate = DateTime.Today;  // Legacy: DTVchr = GetServerDate
        private string exchDiffAccNo = string.Empty;
        private bool authorizeEnabled = true;

        private List<CustomerLookupModel> customers = new();
        private List<string> countries = new();
        private List<CustomInvoiceAuthItemModel> allItems = new();

        // Grid sorting (legacy LV_ColumnClick)
        private string sortColumn = "DT";
        private bool sortAscending = true;

        protected override async Task OnInitializedAsync()
        {
            // Legacy Form_Load
            customers = (await CustomInvoiceAuthService.GetCustomersAsync()).ToList();
            countries = (await CustomInvoiceAuthService.GetCountriesAsync(null)).ToList();

            salesAccounts = (await CustomInvoiceAuthService.GetSalesAccountsAsync()).ToList();
            if (salesAccounts.Count > 0)
                selectedSalesAccount = salesAccounts[0];    // Legacy: If CmbAccTitle.ListCount > 0 Then CmbAccTitle.ListIndex = 0

            exchDiffAccNo = await CustomInvoiceAuthService.GetExchDiffAccNoAsync();
            if (string.IsNullOrEmpty(exchDiffAccNo))
                authorizeEnabled = false;   // Legacy: MsgBox "Exch. Diff. Account not found in general data.", vbCritical / cmdUpdate.Enabled = False

            await LoadData();
        }

        private async Task LoadData()
        {
            isLoading = true;
            StateHasChanged();

            // Legacy RefreshLV
            allItems = (await CustomInvoiceAuthService.GetPendingInvoicesAsync(
                customerCode: selectedCustomer?.CustCode ?? "0",
                country: selectedCountry,
                fromDate: fromDate,
                toDate: toDate
            )).ToList();

            sortColumn = "DT";
            sortAscending = true;
            isLoading = false;
            StateHasChanged();
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            return await Task.FromResult(customers.Where(x => x.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task OnCustomerChanged(CustomerLookupModel? customer)
        {
            selectedCustomer = customer;
            // Legacy cmbCust_matched: refill the country list for the selected customer, back to <All Countries>
            countries = (await CustomInvoiceAuthService.GetCountriesAsync(customer?.CustCode)).ToList();
            selectedCountry = string.Empty;
            await LoadData();
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? string.Empty;
            await LoadData();
        }

        // Legacy cmbRange_Click: DtTo = GetServerDate; presets disable the date pickers, "Date Range" enables them
        private async Task OnDateRangeChanged()
        {
            if (selectedDateRange != "custom")
            {
                toDate = DateTime.Today;
                fromDate = selectedDateRange == "0"
                    ? DateTime.Today    // Todays
                    : DateTime.Today.AddDays(-int.Parse(selectedDateRange));
                await LoadData();
            }
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchSalesAccounts(string searchText)
        {
            return await Task.FromResult(salesAccounts.Where(x => x.AccTitle != null && x.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        private void OnSalesAccountChanged(ChartOfAccountsModel? account)
        {
            selectedSalesAccount = account;
        }

        // Legacy LV_ItemCheck: checking a row selects it and opens the exchange-rate editor
        private void OnRowChecked(CustomInvoiceAuthItemModel item, ChangeEventArgs e)
        {
            item.IsChecked = e.Value is bool isChecked && isChecked;
            if (item.IsChecked)
                BeginRateEdit(item);
        }

        // Legacy LV_DblClick: open the inline exchange-rate editor (txtDesc) for the row
        private void BeginRateEdit(CustomInvoiceAuthItemModel item)
        {
            foreach (var other in allItems.Where(x => x.IsEditingRate && x != item))
                other.IsEditingRate = false;

            item.EditingExchRate = item.ExchRate;
            item.IsEditingRate = true;
        }

        // Legacy txtDesc_KeyPress (Enter): validate, set Exch. Rate and compute Amt PK = Round(Rate * Invoice Amt)
        private void CommitRateEdit(CustomInvoiceAuthItemModel item)
        {
            if (item.EditingExchRate <= 0)
            {
                NotificationServiceManager.ShowError("Invalid Exchange Rate.", "Please enter a valid exchange rate greater than zero.");
                return; // Legacy keeps the editor open on an invalid rate
            }
            item.ExchRate = item.EditingExchRate;
            item.AmtInPakRs = Math.Round(item.ExchRate * item.TotalCustomAmt);
            item.IsEditingRate = false;
        }

        private void OnRateKeyDown(KeyboardEventArgs e, CustomInvoiceAuthItemModel item)
        {
            if (e.Key == "Enter")
                CommitRateEdit(item);
            else if (e.Key == "Escape")
                item.IsEditingRate = false; // Legacy txtDesc_KeyPress (Esc): hide the editor without committing
        }

        private void SortBy(string column)
        {
            if (sortColumn == column)
                sortAscending = !sortAscending;
            else
            {
                sortColumn = column;
                sortAscending = true;
            }
        }

        private string GetSortIcon(string column)
        {
            if (sortColumn != column) return "fa-sort text-muted opacity-50";
            return sortAscending ? "fa-sort-up" : "fa-sort-down";
        }

        private IEnumerable<CustomInvoiceAuthItemModel> SortedItems
        {
            get
            {
                return (sortColumn, sortAscending) switch
                {
                    ("Customer", true) => allItems.OrderBy(x => x.CustCode),
                    ("Customer", false) => allItems.OrderByDescending(x => x.CustCode),
                    ("Country", true) => allItems.OrderBy(x => x.Country),
                    ("Country", false) => allItems.OrderByDescending(x => x.Country),
                    ("Invoice", true) => allItems.OrderBy(x => x.CustomInvoice),
                    ("Invoice", false) => allItems.OrderByDescending(x => x.CustomInvoice),
                    ("DT", false) => allItems.OrderByDescending(x => x.DT),
                    ("Amt", true) => allItems.OrderBy(x => x.TotalCustomAmt),
                    ("Amt", false) => allItems.OrderByDescending(x => x.TotalCustomAmt),
                    ("Rate", true) => allItems.OrderBy(x => x.ExchRate),
                    ("Rate", false) => allItems.OrderByDescending(x => x.ExchRate),
                    ("AmtPK", true) => allItems.OrderBy(x => x.AmtInPakRs),
                    ("AmtPK", false) => allItems.OrderByDescending(x => x.AmtInPakRs),
                    ("SNo", false) => allItems.AsEnumerable().Reverse(),
                    _ => allItems.OrderBy(x => x.DT)
                };
            }
        }

        // Legacy cmdUpdate_Click
        private async Task AuthorizeSelectedInvoices()
        {
            if (isAuthorizing) return;

            if (selectedSalesAccount == null)
            {
                NotificationServiceManager.ShowWarning("Select Account No.", "Please select a Sales Account before authorizing.");
                return;
            }

            var checkedItems = allItems.Where(x => x.IsChecked).ToList();
            if (checkedItems.Count == 0)
            {
                NotificationServiceManager.ShowWarning("No Invoices Selected", "Please check at least one invoice to authorize.");
                return;
            }

            // Guard against the legacy quirk of posting an empty voucher when the exchange rate was never entered
            var invalidItems = checkedItems.Where(x => x.ExchRate <= 0 || x.AmtInPakRs <= 0).ToList();
            if (invalidItems.Count > 0)
            {
                NotificationServiceManager.ShowError("Invalid Exchange Rate.",
                    $"Please enter a valid exchange rate for invoice(s): {string.Join(", ", invalidItems.Select(x => x.CustomInvoice))}");
                return;
            }

            try
            {
                isAuthorizing = true;
                int authorizedCount = await CustomInvoiceAuthService.AuthorizeInvoicesAsync(
                    voucherDate, selectedSalesAccount.AccNo, exchDiffAccNo, checkedItems);

                NotificationServiceManager.ShowSuccess("Authorization Complete", $"{authorizedCount} invoice(s) authorized successfully.");
                await LoadData();   // Legacy: Call RefreshLV
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Authorization Error", ex.Message);
            }
            finally
            {
                isAuthorizing = false;
            }
        }

        // Legacy cmdClose_Click: Unload Me
        private void Close()
        {
            NavigationManager.NavigateTo("/financial");
        }
    }
}
