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

namespace Impulse.Pages.Export.Orders
{
    public partial class OrderItemList : ComponentBase
    {
        [Inject] private IOrderItemListService OrderItemListService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = null!;

        // ─── Filter State ─────────────────────────────────────────────────────────
        private CustomerLookupModel? selectedCustomer = null;
        private string? filterCountry = null;
        private OrderItemGroupLookupModel? selectedItemGroup = null;
        private OrderNoLookupModel? selectedOrderNo = null;
        private bool filterNoDelivery = false;
        private bool filterBoth = false;
        private bool filterUrgent = false;
        private int filterRange = 0; // 0=Next15 1=Next30 2=Next45 3=Custom
        private DateTime dtFrom = DateTime.Today;
        private DateTime dtTo = DateTime.Today.AddDays(15);
        private bool isCustomRange = false;

        // ─── Lookup Data ──────────────────────────────────────────────────────────
        private List<CustomerLookupModel> customers = new();
        private List<string> countries = new();
        private List<OrderItemGroupLookupModel> itemGroups = new();
        private List<OrderNoLookupModel> orderNos = new();

        // ─── Grid Data ────────────────────────────────────────────────────────────
        private List<OrderItemListViewModel> allItems = new();
        private bool isLoading = true;

        // ─── Inline Delivery Edit Modal ────────────────────────────────────────────
        private bool showEditModal = false;
        private DeliveryEditModel editModel = new();
        private bool isEditAuthorized = false;  // stays true once authorized once per session

        // ─── Password Modal ───────────────────────────────────────────────────────
        private bool showPasswordModal = false;
        private string enteredPassword = string.Empty;
        private string passwordError = string.Empty;
        private Func<Task>? postPasswordAction = null;

        // ─── User Info ────────────────────────────────────────────────────────────
        private string currentUserName = "System";
        private string machineName = Environment.MachineName;

        // ─────────────────────────────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            // Grab logged-in username
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            currentUserName = authState.User.Identity?.Name ?? "System";

            // Seed lookup data
            customers = await CustomerOrderService.GetCustomersAsync();
            itemGroups = await OrderItemListService.GetItemGroupsAsync();
            countries = await OrderItemListService.GetCountriesForCustomerAsync(string.Empty);
            orderNos = await OrderItemListService.GetOrderNosForCustomerAsync(string.Empty);

            // Set initial date range (Next 15 days like legacy default cmbRange.ListIndex = 0)
            ApplyDateRange(0);

            await LoadData();
        }

        // ─── Filter Helpers ───────────────────────────────────────────────────────
        private void ApplyDateRange(int rangeIndex)
        {
            filterRange = rangeIndex;
            dtFrom = DateTime.Today;
            isCustomRange = false;
            switch (rangeIndex)
            {
                case 0: dtTo = DateTime.Today.AddDays(15); break;
                case 1: dtTo = DateTime.Today.AddDays(30); break;
                case 2: dtTo = DateTime.Today.AddDays(45); break;
                case 3: isCustomRange = true; break;
            }
        }

        private async Task OnDateRangeChanged(ChangeEventArgs e)
        {
            ApplyDateRange(int.Parse(e.Value?.ToString() ?? "0"));
            await LoadData();
        }

        public async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return customers;
            return await Task.FromResult(customers.Where(c => c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public async Task<IEnumerable<OrderItemGroupLookupModel>> SearchItemGroups(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return itemGroups;
            return await Task.FromResult(itemGroups.Where(g => g.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public async Task<IEnumerable<OrderNoLookupModel>> SearchOrderNos(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return orderNos;
            return await Task.FromResult(orderNos.Where(o => o.DisplayText.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            selectedCustomer = cust;
            filterCountry = null;
            selectedOrderNo = null;
            string custCodeStr = selectedCustomer?.CustCode ?? string.Empty;
            countries = await OrderItemListService.GetCountriesForCustomerAsync(custCodeStr);
            orderNos = await OrderItemListService.GetOrderNosForCustomerAsync(custCodeStr);
            await LoadData();
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            var val = e.Value?.ToString();
            filterCountry = (val == "<All Countries>") ? null : val;
            await LoadData();
        }

        private async Task OnGroupChanged(OrderItemGroupLookupModel? group)
        {
            selectedItemGroup = group;
            await LoadData();
        }

        private async Task OnOrderNoChanged(OrderNoLookupModel? orderNo)
        {
            selectedOrderNo = orderNo;
            await LoadData();
        }

        private async Task OnUrgentChanged(ChangeEventArgs e)
        {
            filterUrgent = (bool)(e.Value ?? false);
            if (filterUrgent) { filterNoDelivery = false; filterBoth = false; }
            await LoadData();
        }

        private async Task OnNoDeliveryChanged(ChangeEventArgs e)
        {
            filterNoDelivery = (bool)(e.Value ?? false);
            if (filterNoDelivery) { filterUrgent = false; filterBoth = false; }
            await LoadData();
        }

        private async Task OnBothChanged(ChangeEventArgs e)
        {
            filterBoth = (bool)(e.Value ?? false);
            if (filterBoth) { filterUrgent = false; filterNoDelivery = false; }
            await LoadData();
        }

        private async Task LoadData()
        {
            isLoading = true;
            StateHasChanged();
            try
            {
                allItems = await OrderItemListService.GetOrderItemsAsync(
                    dtFrom, dtTo,
                    selectedCustomer?.CustCode, filterCountry,
                    selectedItemGroup?.ID, selectedOrderNo?.OrderNo,
                    filterNoDelivery, filterBoth, filterUrgent);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Load Error", ex.Message);
            }
            finally
            {
                isLoading = false;
                StateHasChanged();
            }
        }

        // ─── Inline Delivery Date Edit ─────────────────────────────────────────────
        private async Task OnRowDoubleClick(OrderItemListViewModel item)
        {
            if (!isEditAuthorized)
            {
                await CheckPasswordThenExecute("EditDeliveryDT", async () =>
                {
                    isEditAuthorized = true;
                    OpenEditModal(item);
                    await Task.CompletedTask;
                });
            }
            else
            {
                OpenEditModal(item);
            }
        }

        private void OpenEditModal(OrderItemListViewModel item)
        {
            editModel = new DeliveryEditModel
            {
                ID = item.ID,
                NewDeliveryDT = item.DeliveryStatus == 0 ? DateTime.Today : item.DeliveryDT,
                Remarks = string.Empty,
                PreviousDeliveryStatus = item.DeliveryStatus,
                PreviousDeliveryDT = item.DeliveryStatus == 0 ? (DateTime?)null : item.DeliveryDT
            };
            showEditModal = true;
        }

        private async Task SaveDeliveryDate()
        {
            try
            {
                bool ok = await OrderItemListService.UpdateDeliveryDateAsync(
                    editModel.ID,
                    editModel.NewDeliveryDT,
                    editModel.PreviousDeliveryStatus,
                    editModel.PreviousDeliveryDT,
                    editModel.Remarks,
                    currentUserName,
                    machineName);

                if (ok)
                {
                    showEditModal = false;
                    NotificationServiceManager.ShowSuccess("Saved", "Delivery date updated successfully.");
                    await LoadData();
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Save Failed", ex.Message);
            }
        }

        // ─── Password Modal ───────────────────────────────────────────────────────
        private async Task CheckPasswordThenExecute(string settingName, Func<Task> action)
        {
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
                string correctPassword = await DbHelper.getDatabasePasswordAsync("EditDeliveryDT");
                if (enteredPassword == correctPassword)
                {
                    showPasswordModal = false;
                    if (postPasswordAction != null)
                        await postPasswordAction.Invoke();
                }
                else
                {
                    passwordError = "Incorrect password. Please try again.";
                }
            }
            catch (Exception ex)
            {
                passwordError = $"Error: {ex.Message}";
            }
        }

        // ─── Reports ──────────────────────────────────────────────────────────────
        private async Task PrintOrderItemList()
        {
            try
            {
                var selectionFormula = BuildSelectionFormula();
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Export.OrderItemList,
                    SelectionFormula = selectionFormula
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Report Error", ex.Message);
            }
        }

        private async Task OpenOptionsBarMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private string BuildSelectionFormula()
        {
            string sel;
            if (filterUrgent)
                sel = "({VFOrderItems.DeliveryStatus}=2 OR {VFOrderItems.Remarks}<>'')";
            else if (filterBoth)
                sel = $"({{VFOrderItems.DeliveryStatus}}=0 OR ({{VFOrderItems.DeliveryDT}}=#{dtFrom:d/M/yyyy}# TO #{dtTo:d/M/yyyy}#))";
            else if (filterNoDelivery)
                sel = "{VFOrderItems.DeliveryStatus}=0";
            else
                sel = $"{{VFOrderItems.DeliveryDT}}=#{dtFrom:d/M/yyyy}# TO #{dtTo:d/M/yyyy}#";

            if (!string.IsNullOrEmpty(selectedCustomer?.CustCode))
                sel += $" AND {{VFOrderItems.CustCode}}='{selectedCustomer.CustCode}'";
            if (!string.IsNullOrEmpty(filterCountry))
                sel += $" AND {{VFOrderItems.Country}}='{filterCountry}'";
            if (selectedItemGroup?.ID > 0)
                sel += $" AND {{VFOrderItems.GroupID}}={selectedItemGroup.ID}";

            sel += " AND ({VFOrderItems.Qty}>{VFOrderItems.ShippedQty})";
            return sel;
        }
    }
}
