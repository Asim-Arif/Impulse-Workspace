using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;

namespace Impulse.Components.Export
{
    public partial class CustomerOrdersSelectionModal : ComponentBase
    {
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = null!;
        [Inject] private INotificationService NotificationService { get; set; } = null!;

        private bool IsVisible { get; set; }
        private bool IsLoadingOrders { get; set; }
        private string OrderSearchText { get; set; } = string.Empty;

        private List<CustomerLookupModel> Customers { get; set; } = new();
        private CustomerLookupModel? SelectedCustomer { get; set; }

        private List<string> Countries { get; set; } = new();
        private string SelectedCountry { get; set; } = string.Empty;

        private List<UnshippedOrderOptionModel> Orders { get; set; } = new();
        private HashSet<string> SelectedOrderNos { get; set; } = new(StringComparer.OrdinalIgnoreCase);

        private TaskCompletionSource<string?>? _tcs;

        private IEnumerable<UnshippedOrderOptionModel> FilteredOrders =>
            string.IsNullOrWhiteSpace(OrderSearchText)
                ? Orders
                : Orders.Where(o =>
                    o.OrderNo.Contains(OrderSearchText, StringComparison.OrdinalIgnoreCase) ||
                    o.InternalRefNo.Contains(OrderSearchText, StringComparison.OrdinalIgnoreCase));

        public async Task<string?> ShowAsync(CustomerLookupModel? preselectedCustomer = null, string? preselectedCountry = null)
        {
            IsVisible = true;
            OrderSearchText = string.Empty;
            SelectedOrderNos.Clear();
            Orders.Clear();
            Countries.Clear();
            SelectedCountry = string.Empty;
            _tcs = new TaskCompletionSource<string?>();

            if (!Customers.Any())
            {
                Customers = await CustomerOrderService.GetCustomersAsync();
            }

            if (preselectedCustomer != null && !string.IsNullOrWhiteSpace(preselectedCustomer.CustCode) && preselectedCustomer.CustCode != "0")
            {
                SelectedCustomer = Customers.FirstOrDefault(c => c.CustCode.Equals(preselectedCustomer.CustCode, StringComparison.OrdinalIgnoreCase))
                    ?? preselectedCustomer;
                await LoadCountriesAndOrdersAsync(preselectedCountry);
            }
            else
            {
                SelectedCustomer = null;
            }

            StateHasChanged();
            return await _tcs.Task;
        }

        private Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<CustomerLookupModel>>(Customers);

            return Task.FromResult<IEnumerable<CustomerLookupModel>>(
                Customers.Where(c =>
                    c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnCustomerChanged(CustomerLookupModel? customer)
        {
            SelectedCustomer = customer;
            await LoadCountriesAndOrdersAsync(null);
        }

        private async Task LoadCountriesAndOrdersAsync(string? preferredCountry)
        {
            Countries.Clear();
            SelectedCountry = string.Empty;
            Orders.Clear();
            SelectedOrderNos.Clear();

            if (SelectedCustomer == null || string.IsNullOrWhiteSpace(SelectedCustomer.CustCode))
                return;

            Countries = await CustomerOrderService.GetCountriesForCustomerAsync(SelectedCustomer.CustCode);
            if (Countries.Any())
            {
                SelectedCountry = !string.IsNullOrWhiteSpace(preferredCountry) &&
                                  preferredCountry != "<All Countries>" &&
                                  Countries.Contains(preferredCountry, StringComparer.OrdinalIgnoreCase)
                    ? Countries.First(c => c.Equals(preferredCountry, StringComparison.OrdinalIgnoreCase))
                    : Countries.First();

                await LoadOrdersAsync();
            }
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            SelectedCountry = e.Value?.ToString() ?? string.Empty;
            await LoadOrdersAsync();
        }

        private async Task LoadOrdersAsync()
        {
            Orders.Clear();
            SelectedOrderNos.Clear();

            if (SelectedCustomer == null || string.IsNullOrWhiteSpace(SelectedCustomer.CustCode) || string.IsNullOrWhiteSpace(SelectedCountry))
                return;

            IsLoadingOrders = true;
            StateHasChanged();
            try
            {
                Orders = await CustomerOrderService.GetUnshippedOrdersForCustomerAsync(SelectedCustomer.CustCode, SelectedCountry);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to load unshipped orders: {ex.Message}");
            }
            finally
            {
                IsLoadingOrders = false;
                StateHasChanged();
            }
        }

        private void ToggleOrder(string orderNo, bool isChecked)
        {
            if (isChecked)
                SelectedOrderNos.Add(orderNo);
            else
                SelectedOrderNos.Remove(orderNo);
        }

        private void ToggleSelectAllFiltered(ChangeEventArgs e)
        {
            bool isChecked = e.Value is bool b && b;
            foreach (var ord in FilteredOrders)
            {
                if (isChecked)
                    SelectedOrderNos.Add(ord.OrderNo);
                else
                    SelectedOrderNos.Remove(ord.OrderNo);
            }
        }

        private void SelectAll()
        {
            foreach (var ord in FilteredOrders)
            {
                SelectedOrderNos.Add(ord.OrderNo);
            }
        }

        private void ClearAll()
        {
            SelectedOrderNos.Clear();
        }

        private void Confirm()
        {
            if (SelectedCustomer == null || string.IsNullOrWhiteSpace(SelectedCustomer.CustCode))
            {
                NotificationService.ShowWarning("Customer Required", "Please select a customer.");
                return;
            }

            if (string.IsNullOrWhiteSpace(SelectedCountry))
            {
                NotificationService.ShowWarning("Country Required", "Please select a country.");
                return;
            }

            if (!SelectedOrderNos.Any())
            {
                NotificationService.ShowWarning("Orders Required", "Please select at least one order.");
                return;
            }

            string ordersCsv = string.Join(",", SelectedOrderNos.Select(o => $"'{o.Replace("'", "''")}'"));
            IsVisible = false;
            _tcs?.TrySetResult(ordersCsv);
        }

        private void Cancel()
        {
            IsVisible = false;
            _tcs?.TrySetResult(null);
        }
    }
}
