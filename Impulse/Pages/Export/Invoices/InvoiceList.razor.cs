using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;

namespace Impulse.Pages.Export.Invoices
{
    public partial class InvoiceList : ComponentBase
    {
        [Parameter] public string InvoiceTypeString { get; set; } = "custom";

        [Inject] private IInvoiceListService InvoiceListService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private IJSRuntime JS { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private int invoiceType = 0; // 0 = Custom, 1 = Commercial, 2 = Bank

        private bool isLoading = true;
        private string searchText = string.Empty;
        private string selectedDateRange = "180";
        private DateTime? fromDate = DateTime.Today.AddDays(-180);
        private DateTime? toDate = DateTime.Today;

        private long selectedCompanyId = 0;
        private string selectedCountry = string.Empty;
        private CustomerLookupModel? selectedCustomer;

        private List<CompanyLookupModel> companies = new();
        private List<CustomerLookupModel> customers = new();
        private List<string> countries = new();
        private List<InvoiceListModel> allItems = new();
        private InvoiceListModel? contextMenuRowItem;

        // Modals
        private bool showPasswordModal = false;
        private string passwordSettingName = string.Empty;
        private string enteredPassword = string.Empty;
        private string passwordError = string.Empty;
        private Func<Task>? postPasswordAction = null;

        protected override async Task OnInitializedAsync()
        {
            await LoadFilters();
        }

        protected override async Task OnParametersSetAsync()
        {
            invoiceType = InvoiceTypeString.ToLower() switch
            {
                "commercial" => 1,
                "bank" => 2,
                _ => 0
            };
            await LoadData();
        }

        public string GetPageTitle()
        {
            return invoiceType switch
            {
                1 => "Commercial Invoice List",
                2 => "Bank Invoice List",
                _ => "Custom Invoice List"
            };
        }

        private async Task LoadFilters()
        {
            companies = (await InvoiceListService.GetCompaniesAsync()).ToList();
            customers = (await InvoiceListService.GetCustomersAsync()).ToList();
            countries = (await InvoiceListService.GetCountriesAsync()).ToList();
        }

        private async Task LoadData()
        {
            isLoading = true;
            StateHasChanged();

            allItems = (await InvoiceListService.GetInvoicesAsync(
                invoiceType: invoiceType,
                companyId: selectedCompanyId,
                customerCode: selectedCustomer?.CustCode ?? "0",
                country: selectedCountry,
                dateRange: selectedDateRange,
                fromDate: fromDate,
                toDate: toDate,
                searchText: searchText
            )).ToList();

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
            await LoadData();
        }

        private async Task OnCompanyChanged(ChangeEventArgs e)
        {
            if (long.TryParse(e.Value?.ToString(), out long compId))
            {
                selectedCompanyId = compId;
                await LoadData();
            }
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? string.Empty;
            await LoadData();
        }

        private async Task OnDateRangeChanged()
        {
            if (selectedDateRange != "custom")
            {
                if (int.TryParse(selectedDateRange, out int days))
                {
                    fromDate = DateTime.Today.AddDays(-days);
                    toDate = DateTime.Today;
                }
                await LoadData();
            }
        }

        private async Task OpenRowContextMenu(MouseEventArgs e, InvoiceListModel item)
        {
            contextMenuRowItem = item;
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY);
        }

        private async Task CheckPasswordThenExecute(string settingName, Func<Task> action)
        {
            passwordSettingName = settingName;
            enteredPassword = string.Empty;
            passwordError = string.Empty;
            postPasswordAction = action;
            showPasswordModal = true;
            await Task.CompletedTask;
        }

        private async Task VerifyPasswordAndExecute()
        {
            try
            {
                string correctPassword = await DbHelper.getDatabasePasswordAsync(passwordSettingName);
                if (enteredPassword == correctPassword)
                {
                    showPasswordModal = false;
                    if (postPasswordAction != null)
                    {
                        await postPasswordAction.Invoke();
                    }
                }
                else
                {
                    passwordError = "Incorrect password. Please try again.";
                }
            }
            catch (Exception ex)
            {
                passwordError = $"Error validating password: {ex.Message}";
            }
        }

        private async Task DeleteInvoice(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;

            bool confirm = await JS.InvokeAsync<bool>("confirm", $"Are you sure you want to delete invoice {contextMenuRowItem.CustomInvoice}?");
            if (!confirm) return;

            _ = CheckPasswordThenExecute("DeleteExportEntries", async () =>
            {
                try
                {
                    await InvoiceListService.DeleteInvoiceAsync(contextMenuRowItem.CustomInvoice);
                    NotificationServiceManager.ShowSuccess("Success", "Invoice deleted successfully.");
                    await LoadData();
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Delete Error", ex.Message);
                }
            });
        }

        private void PrintReport(string reportName)
        {
            if (contextMenuRowItem == null) return;
            
            var request = new ReportRequest
            {
                ReportName = reportName,
                SelectionFormula = $"{{VCustomInvoiceList.CustomInvoice}}='{contextMenuRowItem.CustomInvoice}'"
            };

            _ = ReportNavigation.PrintReportAsync(request);
        }

        private void AddInvoice(ItemClickEventArgs e)
        {
            // Placeholder: NavigationManager.NavigateTo($"/Export/Invoices/InvoiceDetail/{invoiceType}");
            NotificationServiceManager.ShowWarning("Pending", "Navigation to new Invoice Detail page will be implemented soon.");
        }

        private void EditInvoice(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            // Placeholder: NavigationManager.NavigateTo($"/Export/Invoices/InvoiceDetail/{invoiceType}/{contextMenuRowItem.CustomInvoice}");
            NotificationServiceManager.ShowWarning("Pending", $"Navigation to Edit Invoice {contextMenuRowItem.CustomInvoice} will be implemented soon.");
        }

        private void UpdateRebateRates(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            NotificationServiceManager.ShowWarning("Pending", "Update Rebate Rates functionality will be implemented soon.");
        }

        private void GenerateBatches(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            NotificationServiceManager.ShowWarning("Pending", "Generate Batches functionality will be implemented soon.");
        }
    }
}
