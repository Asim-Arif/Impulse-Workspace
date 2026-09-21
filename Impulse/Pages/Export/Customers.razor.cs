using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.AspNetCore.Components;
using Impulse.Services;

namespace Impulse.Pages.Export
{
    public partial class Customers : ComponentBase
    {
        [Inject] private IDBHelperService _dbHelper { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private NavigationManager Navigation { get; set; } = null!;

        private bool isLoading = true;
        private string searchText = string.Empty;
        private string selectedCountryFilter = "All";
        private string selectedStatusFilter = "All";
        private string selectedCurrencyFilter = "All";

        private List<ForeignCustomerListViewModel> allCustomers = new();
        private List<string> countries = new();
        private List<string> currencies = new();
        private ForeignCustomerListViewModel? customerToDelete = null;
        private ForeignCustomerListViewModel? highlightedCustomer = null;

        private int TotalCount => allCustomers.Count;
        private int ActiveCount => allCustomers.Count(c => c.Active);
        private int InactiveCount => allCustomers.Count(c => !c.Active);
        private int CountryCount => countries.Count;

        private bool HasActiveFilters =>
            !string.IsNullOrWhiteSpace(searchText) ||
            selectedCountryFilter != "All" ||
            selectedStatusFilter != "All" ||
            selectedCurrencyFilter != "All";

        private IEnumerable<ForeignCustomerListViewModel> FilteredCustomers
        {
            get
            {
                var query = allCustomers.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    var search = searchText.Trim();
                    query = query.Where(c =>
                        (c.Name?.Contains(search, StringComparison.OrdinalIgnoreCase) == true) ||
                        (c.CustCode?.Contains(search, StringComparison.OrdinalIgnoreCase) == true) ||
                        (c.Country?.Contains(search, StringComparison.OrdinalIgnoreCase) == true) ||
                        (c.City?.Contains(search, StringComparison.OrdinalIgnoreCase) == true) ||
                        (c.Cont1name?.Contains(search, StringComparison.OrdinalIgnoreCase) == true) ||
                        (c.Email1?.Contains(search, StringComparison.OrdinalIgnoreCase) == true) ||
                        (c.Phone1?.Contains(search, StringComparison.OrdinalIgnoreCase) == true));
                }

                if (!string.IsNullOrWhiteSpace(selectedCountryFilter) && selectedCountryFilter != "All")
                {
                    query = query.Where(c => string.Equals(c.Country, selectedCountryFilter, StringComparison.OrdinalIgnoreCase));
                }

                if (!string.IsNullOrWhiteSpace(selectedCurrencyFilter) && selectedCurrencyFilter != "All")
                {
                    query = query.Where(c => string.Equals(c.Curr, selectedCurrencyFilter, StringComparison.OrdinalIgnoreCase));
                }

                if (selectedStatusFilter == "Active")
                {
                    query = query.Where(c => c.Active);
                }
                else if (selectedStatusFilter == "Inactive")
                {
                    query = query.Where(c => !c.Active);
                }

                return query;
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadCustomers();
        }

        private async Task LoadCustomers()
        {
            isLoading = true;
            try
            {
                allCustomers = await _dbHelper.GetListAsync<ForeignCustomerListViewModel>(
                    "CustCode, Name, City, Country, Curr, Active, Phone1, Email1, Cont1name",
                    "VFCustomers",
                    "ORDER BY Name");

                countries = allCustomers
                    .Select(c => c.Country?.Trim() ?? "")
                    .Where(c => !string.IsNullOrEmpty(c))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(c => c)
                    .ToList();

                currencies = allCustomers
                    .Select(c => c.Curr?.Trim() ?? "")
                    .Where(c => !string.IsNullOrEmpty(c))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(c => c)
                    .ToList();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private void HighlightCustomer(ForeignCustomerListViewModel customer)
        {
            highlightedCustomer = (highlightedCustomer?.CustCode == customer.CustCode && highlightedCustomer?.Country == customer.Country)
                ? null
                : customer;
        }

        private void EditCustomer(ForeignCustomerListViewModel customer)
        {
            Navigation.NavigateTo($"/export/customer/{customer.CustCode + customer.Country}?returnUrl=/customers");
        }

        private void CreateQuotation(ForeignCustomerListViewModel customer)
        {
            Navigation.NavigateTo($"/export/customer-quotation-list?custId={Uri.EscapeDataString(customer.CustCode)}");
        }

        private void CreateOrder(ForeignCustomerListViewModel customer)
        {
            Navigation.NavigateTo($"/export/customer-order-list?custId={Uri.EscapeDataString(customer.CustCode)}");
        }

        private void ClearFilters()
        {
            searchText = string.Empty;
            selectedCountryFilter = "All";
            selectedStatusFilter = "All";
            selectedCurrencyFilter = "All";
        }

        private void OnCountryFilterChanged(ChangeEventArgs e)
        {
            selectedCountryFilter = e.Value?.ToString() ?? "All";
        }

        private void OnStatusFilterChanged(ChangeEventArgs e)
        {
            selectedStatusFilter = e.Value?.ToString() ?? "All";
        }

        private void OnCurrencyFilterChanged(ChangeEventArgs e)
        {
            selectedCurrencyFilter = e.Value?.ToString() ?? "All";
        }

        private void ConfirmDelete(ForeignCustomerListViewModel customer)
        {
            customerToDelete = customer;
        }

        private void CancelDelete()
        {
            customerToDelete = null;
        }

        private async Task DeleteCustomer()
        {
            if (customerToDelete == null) return;

            try
            {
                await _dbHelper.ExecuteAsync(
                    "DELETE FROM ForeignCustomers WHERE CustCode = @CustCode AND Country = @Country",
                    new { customerToDelete.CustCode, customerToDelete.Country });

                NotificationServiceManager.ShowSuccess("Deleted",
                    $"Customer '{customerToDelete.Name}' has been deleted.");
                customerToDelete = null;
                await LoadCustomers();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
                customerToDelete = null;
            }
        }
    }
}
