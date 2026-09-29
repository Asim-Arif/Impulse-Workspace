using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Production.OrderLotsTracking
{
    public partial class OrderLotsTracking : ComponentBase
    {
        [Inject] public IOrderLotsTrackingService TrackingService { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;

        [Parameter]
        [SupplyParameterFromQuery(Name = "orderNo")]
        public string? QueryOrderNo { get; set; }

        [Parameter]
        [SupplyParameterFromQuery(Name = "custCode")]
        public string? QueryCustCode { get; set; }

        public OrderLotsTrackingFilter Filter { get; set; } = new OrderLotsTrackingFilter();
        public OrderLotsTrackingDashboardDto DashboardData { get; set; } = new OrderLotsTrackingDashboardDto();

        public List<TrackingCustomerLookupItem> Customers { get; set; } = new();
        public List<TrackingOrderLookupItem> Orders { get; set; } = new();
        public List<string> HubNames { get; set; } = new();

        public TrackingCustomerLookupItem? SelectedCustomer { get; set; }
        public TrackingOrderLookupItem? SelectedOrder { get; set; }

        public bool IsLoading { get; set; } = false;
        public string ClientSearchTerm { get; set; } = string.Empty;

        public List<OrderLotTrackingItemDto> FilteredLots
        {
            get
            {
                var list = DashboardData.Lots ?? new List<OrderLotTrackingItemDto>();
                if (string.IsNullOrWhiteSpace(ClientSearchTerm))
                    return list;

                string term = ClientSearchTerm.Trim().ToLowerInvariant();
                return list.Where(l =>
                    (l.LotNo != null && l.LotNo.ToLowerInvariant().Contains(term)) ||
                    (l.OrderNo != null && l.OrderNo.ToLowerInvariant().Contains(term)) ||
                    (l.CustCode != null && l.CustCode.ToLowerInvariant().Contains(term)) ||
                    (l.CustomerName != null && l.CustomerName.ToLowerInvariant().Contains(term)) ||
                    (l.ItemCode != null && l.ItemCode.ToLowerInvariant().Contains(term)) ||
                    (l.ItemName != null && l.ItemName.ToLowerInvariant().Contains(term)) ||
                    (l.ProcessName != null && l.ProcessName.ToLowerInvariant().Contains(term)) ||
                    (l.Hub_Name != null && l.Hub_Name.ToLowerInvariant().Contains(term)) ||
                    (l.Supervisors != null && l.Supervisors.ToLowerInvariant().Contains(term)) ||
                    (l.MakerName != null && l.MakerName.ToLowerInvariant().Contains(term)) ||
                    (l.TargetStatus != null && l.TargetStatus.ToLowerInvariant().Contains(term))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoading = true;

                // Load filter lookups
                var customerTask = TrackingService.GetCustomersLookupAsync();
                var hubTask = TrackingService.GetHubNamesLookupAsync();

                await Task.WhenAll(customerTask, hubTask);

                Customers = await customerTask;
                HubNames = await hubTask;

                // Pre-populate from query string if available
                if (!string.IsNullOrWhiteSpace(QueryCustCode))
                {
                    Filter.CustCode = QueryCustCode.Trim();
                    SelectedCustomer = Customers.FirstOrDefault(c => c.CustCode.Equals(Filter.CustCode, StringComparison.OrdinalIgnoreCase));
                }

                if (!string.IsNullOrWhiteSpace(QueryOrderNo))
                {
                    Filter.OrderNo = QueryOrderNo.Trim();
                    Filter.DateRangeType = 0; // All time if direct order queried
                }

                // Load orders for selected customer or all recent orders
                Orders = await TrackingService.GetOrdersLookupAsync(Filter.CustCode);
                if (!string.IsNullOrWhiteSpace(Filter.OrderNo))
                {
                    SelectedOrder = Orders.FirstOrDefault(o => o.OrderNo.Equals(Filter.OrderNo, StringComparison.OrdinalIgnoreCase));
                }

                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Initialization Error", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true;
                DashboardData = await TrackingService.GetOrderLotsTrackingAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Data Error", $"Failed to load lots: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task OnCustomerSelected(TrackingCustomerLookupItem? cust)
        {
            SelectedCustomer = cust;
            Filter.CustCode = cust?.CustCode;

            // Reload orders for this customer
            SelectedOrder = null;
            Filter.OrderNo = null;
            Orders = await TrackingService.GetOrdersLookupAsync(Filter.CustCode);

            await LoadDataAsync();
        }

        public async Task OnOrderSelected(TrackingOrderLookupItem? order)
        {
            SelectedOrder = order;
            Filter.OrderNo = order?.OrderNo;
            await LoadDataAsync();
        }

        public async Task OnHubFilterChanged(string? hubName)
        {
            Filter.HubName = hubName;
            await LoadDataAsync();
        }

        public async Task ToggleHubCardFilter(string hubName)
        {
            if (string.Equals(Filter.HubName, hubName, StringComparison.OrdinalIgnoreCase))
            {
                Filter.HubName = null;
            }
            else
            {
                Filter.HubName = hubName;
            }
            await LoadDataAsync();
        }

        public async Task OnDateRangeChanged()
        {
            await LoadDataAsync();
        }

        public async Task OnDateFilterModeChanged()
        {
            await LoadDataAsync();
        }

        public async Task OnIncludeCompletedChanged()
        {
            await LoadDataAsync();
        }

        public async Task ResetFiltersAsync()
        {
            Filter = new OrderLotsTrackingFilter
            {
                DateRangeType = 3,
                DateFilterMode = 0,
                DtFrom = DateTime.Today.AddDays(-30),
                DtTo = DateTime.Today,
                IncludeCompleted = false
            };
            SelectedCustomer = null;
            SelectedOrder = null;
            ClientSearchTerm = string.Empty;

            Orders = await TrackingService.GetOrdersLookupAsync(null);
            await LoadDataAsync();
        }

        public void NavigateToOrderManagement(string orderNo)
        {
            if (!string.IsNullOrWhiteSpace(orderNo))
            {
                NavigationManager.NavigateTo($"/production/order-management?orderNo={Uri.EscapeDataString(orderNo)}");
            }
        }

        public string GetRowClass(OrderLotTrackingItemDto item)
        {
            if (item.IsCompleted) return "table-light";
            if (item.TargetStatus == "Overdue") return "table-danger-subtle";
            if (item.TargetStatus == "Due Soon" || item.TargetStatus == "Due Today") return "table-warning-subtle";
            return string.Empty;
        }
    }
}
