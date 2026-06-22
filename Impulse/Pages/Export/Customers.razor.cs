using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.AspNetCore.Components;
using Impulse.Services;

namespace Impulse.Pages.Export
{
    public partial class Customers : ComponentBase
    {
        [Inject] private IDBHelperService _dbHelper { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;

        private bool isLoading = true;
        private string searchText = string.Empty;
        private List<ForeignCustomerListViewModel> allCustomers = new();
        private ForeignCustomerListViewModel? customerToDelete = null;

        private IEnumerable<ForeignCustomerListViewModel> FilteredCustomers =>
            string.IsNullOrWhiteSpace(searchText)
                ? allCustomers
                : allCustomers.Where(c =>
                    c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    c.Country.Contains(searchText, StringComparison.OrdinalIgnoreCase));

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
                    "CustCode, Name, City, Country, Curr, Active",
                    "VFCustomers",
                    "ORDER BY Name");
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
