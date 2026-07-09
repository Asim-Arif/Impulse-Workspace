using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components.Authorization;
using System.Security.Claims;
using BlazorContextMenu;
using Microsoft.JSInterop;

namespace Impulse.Pages.Export.Orders
{
    public partial class CustomerItemBalances : ComponentBase
    {
        [Inject] private ICustomerItemBalancesService CustomerItemBalancesService { get; set; } = null!;
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private IJSRuntime JS { get; set; } = null!;

        // ─── State ────────────────────────────────────────────────────────────────
        private List<CustomerLookupModel> customers = new();
        private CustomerLookupModel? selectedCustomer = null;
        private string searchText = string.Empty;

        private List<CustomerItemBalancesViewModel> allItems = new();
        private bool isLoading = false;

        private IEnumerable<CustomerItemBalancesViewModel> FilteredItems
        {
            get
            {
                if (string.IsNullOrWhiteSpace(searchText))
                    return allItems;

                return allItems.Where(i =>
                    i.OrderNo.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    i.InternalRefNo.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    i.CompItemCode.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    i.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                );
            }
        }

        // ─── Auth / Password State ─────────────────────────────────────────────────
        private bool showPasswordModal = false;
        private string enteredPassword = string.Empty;
        private string passwordError = string.Empty;
        private Func<Task>? postPasswordAction = null;
        private bool isEditAuthorized = false;

        protected override async Task OnInitializedAsync()
        {
            customers = await CustomerOrderService.GetCustomersAsync();
        }

        public async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string search)
        {
            if (string.IsNullOrWhiteSpace(search)) return customers;
            return await Task.FromResult(customers.Where(c => c.CustCode.Contains(search, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            selectedCustomer = cust;
            await LoadData();
        }

        private async Task LoadData()
        {
            if (selectedCustomer == null || string.IsNullOrWhiteSpace(selectedCustomer.CustCode))
            {
                allItems.Clear();
                return;
            }

            isLoading = true;
            try
            {
                allItems = await CustomerItemBalancesService.GetCustomerItemBalancesAsync(selectedCustomer.CustCode);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Loading Data", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        // ─── Inline Editing ────────────────────────────────────────────────────────
        private void BeginEditBalQty(CustomerItemBalancesViewModel item)
        {
            item.EditBalQty = item.BalQty;
            item.IsEditing = true;
        }

        private void CancelEdit(CustomerItemBalancesViewModel item)
        {
            item.IsEditing = false;
        }

        private async Task HandleBalQtyKeyUp(KeyboardEventArgs e, CustomerItemBalancesViewModel item)
        {
            if (e.Key == "Escape")
            {
                CancelEdit(item);
                return;
            }

            if (e.Key == "Enter")
            {
                // Validate
                if (item.EditBalQty < 0)
                {
                    NotificationServiceManager.ShowWarning("Validation", "Invalid Qty.");
                    return;
                }

                if (isEditAuthorized)
                {
                    await ExecuteUpdateBalQty(item);
                }
                else
                {
                    PromptPassword(async () => await ExecuteUpdateBalQty(item));
                }
            }
        }

        private async Task ExecuteUpdateBalQty(CustomerItemBalancesViewModel item)
        {
            try
            {
                decimal qtyDiff = item.EditBalQty - item.BalQty;
                if (qtyDiff != 0)
                {
                    await CustomerItemBalancesService.UpdateBalanceQuantityAsync(item.ID, qtyDiff);
                    
                    // Update UI state
                    item.BalQty = item.EditBalQty;
                    item.Qty += qtyDiff;
                    NotificationServiceManager.ShowSuccess("Success", $"Balance Qty updated successfully.");
                }
                item.IsEditing = false;
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Updating", ex.Message);
            }
        }

        private void ToggleDeliveryStatus(CustomerItemBalancesViewModel item)
        {
            if (isEditAuthorized)
            {
                _ = ExecuteToggleDeliveryStatus(item);
            }
            else
            {
                PromptPassword(async () => await ExecuteToggleDeliveryStatus(item));
            }
        }

        private async Task ExecuteToggleDeliveryStatus(CustomerItemBalancesViewModel item)
        {
            try
            {
                // Toggle between Urgent (2) and Regular (1). If No Date (0), keep it? 
                // Legacy: If Urgent (2) -> Regular (1). Else -> Urgent (2).
                int newStatus = item.DeliveryStatus == 2 ? 1 : 2;

                await CustomerItemBalancesService.ToggleDeliveryStatusAsync(item.ID, newStatus);
                item.DeliveryStatus = newStatus;
                
                NotificationServiceManager.ShowSuccess("Success", $"Status updated.");
                StateHasChanged();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Updating", ex.Message);
            }
        }

        // ─── Helpers ───────────────────────────────────────────────────────────────
        private string GetDeliveryStatusText(int status)
        {
            return status switch
            {
                0 => "No Date",
                1 => "Regular",
                2 => "Urgent",
                _ => "Unknown"
            };
        }

        private string GetDeliveryStatusCssClass(int status)
        {
            return status switch
            {
                0 => "bg-secondary",
                1 => "bg-info",
                2 => "bg-danger",
                _ => "bg-light text-dark"
            };
        }

        // ─── Context Menu / Reports ────────────────────────────────────────────────
        private async Task OpenOptionsBarMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private async Task PrintList()
        {
            if (selectedCustomer == null)
            {
                NotificationServiceManager.ShowWarning("Warning", "Please select a customer first.");
                return;
            }
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "OrderItemBalances.rpt",
                    SelectionFormula = $"{{VFOrderItems.CustCode}}='{selectedCustomer.CustCode}' AND ({{VFOrderItems.Qty}}>{{VFOrderItems.ShippedQty}}) AND {{VFOrderItems.CustCode}} NOT IN('Stock')"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Report Error", ex.Message);
            }
        }

        private async Task PrintUrgent()
        {
            if (selectedCustomer == null)
            {
                NotificationServiceManager.ShowWarning("Warning", "Please select a customer first.");
                return;
            }
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "UrgentRequest.rpt",
                    SelectionFormula = $"{{VrptOrders.CustCode}}='{selectedCustomer.CustCode}' AND {{VrptOrders.DeliveryStatus}}=2"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Report Error", ex.Message);
            }
        }

        // ─── Authorization Modal ───────────────────────────────────────────────────
        private void PromptPassword(Func<Task> onSuccess)
        {
            enteredPassword = string.Empty;
            passwordError = string.Empty;
            postPasswordAction = onSuccess;
            showPasswordModal = true;
        }

        private async Task VerifyPasswordAndExecute()
        {
            try
            {
                passwordError = string.Empty;
                var correctPassword = await DbHelper.getDatabasePasswordAsync("EditDeliveryDT");

                if (enteredPassword == correctPassword)
                {
                    isEditAuthorized = true;
                    showPasswordModal = false;
                    if (postPasswordAction != null)
                    {
                        await postPasswordAction.Invoke();
                        postPasswordAction = null;
                    }
                }
                else
                {
                    passwordError = "Incorrect password.";
                }
            }
            catch (Exception ex)
            {
                passwordError = $"Error: {ex.Message}";
            }
        }
    }
}
