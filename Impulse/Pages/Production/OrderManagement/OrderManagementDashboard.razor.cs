using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Production.OrderManagement
{
    public class ChartDataItem
    {
        public string Category { get; set; } = string.Empty;
        public double Value { get; set; }
        public string Color { get; set; } = "#3b82f6";
    }

    public class SvgDonutSlice
    {
        public string Category { get; set; } = string.Empty;
        public double Value { get; set; }
        public string Color { get; set; } = "#3b82f6";
        public double Percentage { get; set; }
        public double DashArray { get; set; }
        public double DashOffset { get; set; }
    }

    public partial class OrderManagementDashboard : ComponentBase
    {
        [Inject] private IOrderManagementService OrderService { get; set; } = default!;
        [Inject] private NavigationManager NavManager { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

        // Filter & KPI Data
        protected OrderManagementFilter Filter { get; set; } = new();
        protected OrderSummaryCardDto Metrics { get; set; } = new();
        protected List<LookupItemString> Customers { get; set; } = new();

        // Chart Data
        protected List<ChartDataItem> OrderStatusChartData { get; set; } = new();
        protected List<ChartDataItem> VolumeChartData { get; set; } = new();
        protected List<SvgDonutSlice> OrderStatusSlices { get; set; } = new();
        protected List<SvgDonutSlice> VolumeSlices { get; set; } = new();

        // Orders & Items
        protected List<CustomerOrderHeaderDto> Orders { get; set; } = new();
        protected CustomerOrderHeaderDto? SelectedOrder { get; set; }
        protected List<OrderItemProgressDto> OrderItems { get; set; } = new();
        protected OrderItemProgressDto? SelectedItem { get; set; }

        // Detail Blocks (for selected item)
        protected List<ItemPurchaseOrderDto> ItemPOs { get; set; } = new();
        protected List<ItemRunningLotDto> ItemLots { get; set; } = new();
        protected List<ItemStockAdjustmentDto> ItemStockAdjustments { get; set; } = new();
        protected List<ItemDispatchDetailDto> ItemDispatches { get; set; } = new();

        // UI States
        protected bool IsLoadingOrders { get; set; } = true;
        protected bool IsLoadingItems { get; set; } = false;
        protected bool IsLoadingDetails { get; set; } = false;
        protected int ActiveTab { get; set; } = 1; // 1 = POs, 2 = Running Lots, 3 = Dispatch Details
        protected bool ShowAllDetailsTogether { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            await LoadInitialDataAsync();
        }

        private async Task LoadInitialDataAsync()
        {
            IsLoadingOrders = true;
            try
            {
                Customers = await OrderService.GetCustomersAsync();
                await RefreshOrdersAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to initialize dashboard: {ex.Message}");
            }
            finally
            {
                IsLoadingOrders = false;
            }
        }

        protected async Task RefreshOrdersAsync()
        {
            IsLoadingOrders = true;
            try
            {
                Orders = await OrderService.GetOrdersAsync(Filter);
                Metrics = await OrderService.GetOrderSummaryMetricsAsync(Filter);
                UpdateChartData();

                // Auto-select first order if none selected or if previous selected is no longer in list
                if (Orders.Any())
                {
                    var orderToSelect = SelectedOrder != null 
                        ? Orders.FirstOrDefault(x => x.OrderNo.Equals(SelectedOrder.OrderNo, StringComparison.OrdinalIgnoreCase)) ?? Orders.First()
                        : Orders.First();

                    await SelectOrderAsync(orderToSelect);
                }
                else
                {
                    SelectedOrder = null;
                    OrderItems.Clear();
                    SelectedItem = null;
                    ClearDetails();
                }
            }
            finally
            {
                IsLoadingOrders = false;
            }
        }

        private void UpdateChartData()
        {
            // 1. Order Status Distribution Chart Data
            var statusList = new List<ChartDataItem>();
            int completed = Orders.Count(x => x.ProgressPct >= 100);
            int inProgress = Orders.Count(x => x.ProgressPct > 0 && x.ProgressPct < 100 && !x.IsOverdue);
            int delayed = Orders.Count(x => x.IsOverdue);
            int planned = Orders.Count(x => x.ProgressPct == 0 && !x.IsOverdue);

            if (completed > 0) statusList.Add(new ChartDataItem { Category = "Completed", Value = completed, Color = "#10b981" });
            if (inProgress > 0) statusList.Add(new ChartDataItem { Category = "In Progress", Value = inProgress, Color = "#3b82f6" });
            if (delayed > 0) statusList.Add(new ChartDataItem { Category = "Delayed", Value = delayed, Color = "#ef4444" });
            if (planned > 0) statusList.Add(new ChartDataItem { Category = "Planned", Value = planned, Color = "#94a3b8" });

            OrderStatusChartData = statusList;
            OrderStatusSlices = CalculateSlices(statusList);

            // 2. Volume Distribution Chart Data (Produced vs Remaining vs Dispatched)
            var volList = new List<ChartDataItem>();
            int dispatched = Metrics.TotalDispatchedPcs;
            int producedWaiting = Math.Max(0, Metrics.TotalProducedPcs - dispatched);
            int remaining = Math.Max(0, Metrics.TotalOrderedPcs - Metrics.TotalProducedPcs);

            if (dispatched > 0) volList.Add(new ChartDataItem { Category = "Dispatched", Value = dispatched, Color = "#10b981" });
            if (producedWaiting > 0) volList.Add(new ChartDataItem { Category = "Ready/WIP", Value = producedWaiting, Color = "#3b82f6" });
            if (remaining > 0) volList.Add(new ChartDataItem { Category = "Pending", Value = remaining, Color = "#f59e0b" });

            VolumeChartData = volList;
            VolumeSlices = CalculateSlices(volList);
        }

        private static List<SvgDonutSlice> CalculateSlices(List<ChartDataItem> items)
        {
            double total = items.Sum(x => x.Value);
            if (total <= 0) return new List<SvgDonutSlice>();

            const double circumference = 276.46; // 2 * PI * 44
            double currentOffset = 0;
            var list = new List<SvgDonutSlice>();

            foreach (var item in items)
            {
                if (item.Value <= 0) continue;
                double pct = (item.Value / total) * 100.0;
                double arcLength = (pct / 100.0) * circumference;

                list.Add(new SvgDonutSlice
                {
                    Category = item.Category,
                    Value = item.Value,
                    Color = item.Color,
                    Percentage = Math.Round(pct, 1),
                    DashArray = arcLength,
                    DashOffset = -currentOffset
                });

                currentOffset += arcLength;
            }

            return list;
        }

        public static string GetProgressBarStyle(double pct)
        {
            if (pct >= 100)
                return "repeating-linear-gradient(-45deg, #059669, #059669 8px, #10b981 8px, #10b981 16px)";
            if (pct < 20)
                return "repeating-linear-gradient(-45deg, #dc2626, #dc2626 8px, #ef4444 8px, #ef4444 16px)";
            return "repeating-linear-gradient(-45deg, #2563eb, #2563eb 8px, #3b82f6 8px, #3b82f6 16px)";
        }

        protected async Task OnCustomerChanged(ChangeEventArgs e)
        {
            Filter.CustCode = e.Value?.ToString() ?? string.Empty;
            await RefreshOrdersAsync();
        }

        protected async Task OnDateRangeTypeChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int rangeType))
            {
                Filter.DateRangeType = rangeType;
                await RefreshOrdersAsync();
            }
        }

        protected async Task SetStatusFilter(string status)
        {
            Filter.StatusFilter = status;
            await RefreshOrdersAsync();
        }

        protected async Task OnSearchKeyDown(Microsoft.AspNetCore.Components.Web.KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await RefreshOrdersAsync();
            }
        }

        protected async Task ResetFiltersAsync()
        {
            Filter = new OrderManagementFilter();
            await RefreshOrdersAsync();
        }

        protected async Task SelectOrderAsync(CustomerOrderHeaderDto order)
        {
            SelectedOrder = order;
            SelectedItem = null;
            ClearDetails();

            IsLoadingItems = true;
            try
            {
                OrderItems = await OrderService.GetOrderItemsAsync(order.OrderNo);

                // Recalculate exact overall order progress from child items if items exist
                if (OrderItems.Any())
                {
                    int totalOrdered = OrderItems.Sum(x => x.OrderedQty);
                    if (totalOrdered > 0)
                    {
                        double weightedProgress = OrderItems.Sum(x => x.OrderedQty * x.ProgressPct) / totalOrdered;
                        order.ProgressPct = Math.Round(weightedProgress, 1);
                    }

                    // Auto-select first item
                    await SelectItemAsync(OrderItems.First());
                }
            }
            finally
            {
                IsLoadingItems = false;
            }
        }

        protected async Task SelectItemAsync(OrderItemProgressDto item)
        {
            SelectedItem = item;
            IsLoadingDetails = true;

            try
            {
                var poTask = OrderService.GetItemPurchaseOrdersAsync(item.OrderNo, item.ItemCode, item.CompItemCode);
                var lotsTask = OrderService.GetItemRunningLotsAsync(item.OrderNo, item.ItemCode, item.CompItemCode);
                var stockTask = OrderService.GetItemStockAdjustmentsAsync(item.OrderNo, item.ItemCode, item.CompItemCode);
                var dispTask = OrderService.GetItemDispatchDetailsAsync(item.OrderNo, item.ItemCode, item.CompItemCode);

                await Task.WhenAll(poTask, lotsTask, stockTask, dispTask);

                ItemPOs = await poTask;
                ItemLots = await lotsTask;
                ItemStockAdjustments = await stockTask;
                ItemDispatches = await dispTask;
            }
            finally
            {
                IsLoadingDetails = false;
            }
        }

        protected void SetActiveTab(int tab)
        {
            ActiveTab = tab;
        }

        protected void ToggleDetailView()
        {
            ShowAllDetailsTogether = !ShowAllDetailsTogether;
        }

        private void ClearDetails()
        {
            ItemPOs.Clear();
            ItemLots.Clear();
            ItemStockAdjustments.Clear();
            ItemDispatches.Clear();
        }
    }
}
