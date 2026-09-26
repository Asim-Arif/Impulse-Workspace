using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.Production.PpcOrderPlanning
{
    public partial class PpcOrderPlanning : ComponentBase
    {
        [Inject] public IPpcOrderPlanningService PpcPlanningService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;
        [Inject] public AuthenticationStateProvider AuthStateProvider { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public string? OrderNoParam { get; set; }
        [SupplyParameterFromQuery(Name = "orderNo")] public string? OrderNoQuery { get; set; }

        public List<PpcOrderSummaryListItem> AuthorizedOrders { get; set; } = new();
        public string SelectedOrderNo { get; set; } = string.Empty;
        public PpcOrderHeaderDto? OrderHeader { get; set; }
        public string PlanningNotes { get; set; } = string.Empty;

        public bool IsLoading { get; set; } = false;
        public bool IsSaving { get; set; } = false;

        public bool CanSavePlan =>
            OrderHeader != null &&
            OrderHeader.Items.Any() &&
            !OrderHeader.Items.Any(i => i.StockQty > i.AvailableInHandStock || i.StockQty < 0) &&
            !OrderHeader.Items.Any(i => i.Purchases.Any(p => p.PurchaseQty <= 0 || p.ProcessID <= 0)) &&
            !OrderHeader.Items.Any(i => i.HubSchedules.Any(h => h.StartDate > h.EndDate));

        protected override async Task OnInitializedAsync()
        {
            await LoadAuthorizedOrdersAsync();

            string? initialOrderNo = !string.IsNullOrWhiteSpace(OrderNoParam)
                ? OrderNoParam
                : OrderNoQuery;

            if (!string.IsNullOrWhiteSpace(initialOrderNo))
            {
                SelectedOrderNo = initialOrderNo.Trim();
                await LoadOrderDetailsAsync(SelectedOrderNo);
            }
        }

        private async Task LoadAuthorizedOrdersAsync()
        {
            try
            {
                AuthorizedOrders = await PpcPlanningService.GetAuthorizedOrdersForPlanningAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to load orders",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task OnOrderSelected(ChangeEventArgs e)
        {
            SelectedOrderNo = e.Value?.ToString() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(SelectedOrderNo))
            {
                await LoadOrderDetailsAsync(SelectedOrderNo);
            }
            else
            {
                OrderHeader = null;
            }
        }

        public async Task ReloadCurrentOrderAsync()
        {
            if (!string.IsNullOrWhiteSpace(SelectedOrderNo))
            {
                await LoadOrderDetailsAsync(SelectedOrderNo);
            }
        }

        private async Task LoadOrderDetailsAsync(string orderNo)
        {
            try
            {
                IsLoading = true;
                OrderHeader = await PpcPlanningService.GetOrderForPlanningAsync(orderNo);
                if (OrderHeader != null)
                {
                    PlanningNotes = OrderHeader.Notes ?? string.Empty;

                    // Ensure Hub Planned Qty is initialized if previously unset
                    foreach (var item in OrderHeader.Items)
                    {
                        foreach (var hub in item.HubSchedules)
                        {
                            if (hub.PlannedQty == 0)
                            {
                                hub.PlannedQty = item.ProductionQty;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Order",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task AddPurchaseSplit(PpcOrderItemDto item)
        {
            var split = new PpcItemPurchaseAllocationDto
            {
                ProcessID = 0,
                ProcessName = "",
                VendID = null,
                MakerName = null,
                PurchaseQty = Math.Max(0, item.ProductionQty),
                PurchaseRate = 0
            };

            // Pre-select first available process if only one exists
            if (item.AvailablePurchaseProcesses.Count == 1)
            {
                var singleProc = item.AvailablePurchaseProcesses.First();
                split.ProcessID = singleProc.ProcessID;
                split.ProcessName = singleProc.ProcessName;
                await LoadMakersForSplitAsync(item, split);
            }

            item.Purchases.Add(split);
            await InvokeAsync(StateHasChanged);
        }

        public void RemovePurchaseSplit(PpcOrderItemDto item, PpcItemPurchaseAllocationDto split)
        {
            item.Purchases.Remove(split);
        }

        public async Task OnPurchaseProcessChanged(PpcOrderItemDto item, PpcItemPurchaseAllocationDto split, ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int procId))
            {
                split.ProcessID = procId;
                var proc = item.AvailablePurchaseProcesses.FirstOrDefault(p => p.ProcessID == procId);
                split.ProcessName = proc?.ProcessName ?? string.Empty;
                split.VendID = null;
                split.MakerName = null;

                await LoadMakersForSplitAsync(item, split);
            }
        }

        private async Task LoadMakersForSplitAsync(PpcOrderItemDto item, PpcItemPurchaseAllocationDto split)
        {
            if (split.ProcessID > 0)
            {
                try
                {
                    split.AvailableMakers = await PpcPlanningService.GetMakersForItemAndProcessAsync(item.CompItemID, split.ProcessID);
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Makers Lookup",
                        Detail = ex.Message,
                        Duration = 3000
                    });
                }
            }
            else
            {
                split.AvailableMakers.Clear();
            }

            await InvokeAsync(StateHasChanged);
        }

        public void OnMakerChanged(PpcItemPurchaseAllocationDto split, ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int vendId) && vendId > 0)
            {
                split.VendID = vendId;
                var maker = split.AvailableMakers.FirstOrDefault(m => m.VendID == vendId);
                split.MakerName = maker?.MakerName;
                split.PurchaseRate = maker?.Rate ?? 0;
            }
            else
            {
                split.VendID = null;
                split.MakerName = null;
                split.PurchaseRate = 0;
            }
        }

        public async Task SavePlanAsync()
        {
            if (OrderHeader == null || !OrderHeader.Items.Any()) return;

            // 1. Validation checks
            foreach (var item in OrderHeader.Items)
            {
                if (item.StockQty > item.AvailableInHandStock)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Invalid Stock Qty",
                        Detail = $"Stock quantity for {item.CompItemID} ({item.StockQty}) exceeds available in-hand stock ({item.AvailableInHandStock}).",
                        Duration = 5000
                    });
                    return;
                }

                if (item.Purchases.Any(p => p.PurchaseQty <= 0 || p.ProcessID <= 0))
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Incomplete Purchase Split",
                        Detail = $"Please select a valid Process and Quantity (> 0) for all purchase allocations in item {item.CompItemID}.",
                        Duration = 5000
                    });
                    return;
                }

                if (item.HubSchedules.Any(h => h.StartDate > h.EndDate))
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Invalid Hub Schedule",
                        Detail = $"Hub Start Date cannot be later than End Date for item {item.CompItemID}.",
                        Duration = 5000
                    });
                    return;
                }

                // Sync hub planned qty if remaining production qty exists
                foreach (var hub in item.HubSchedules)
                {
                    if (hub.PlannedQty == 0 && item.ProductionQty > 0)
                    {
                        hub.PlannedQty = item.ProductionQty;
                    }
                }
            }

            IsSaving = true;

            try
            {
                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                var currentUserName = authState.User.Identity?.Name ?? "PPC Planner";

                var request = new PpcPlanSaveRequest
                {
                    OrderNo = OrderHeader.OrderNo,
                    PlannedBy = currentUserName,
                    Notes = PlanningNotes,
                    Items = OrderHeader.Items
                };

                bool ok = await PpcPlanningService.SavePpcPlanAndDispatchTasksAsync(request);

                if (ok)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "PPC Plan Saved & Dispatched",
                        Detail = $"Plan for Order #{OrderHeader.OrderNo} saved successfully. Tasks dispatched to Stock, Purchaser, and Hub Supervisors.",
                        Duration = 5000
                    });

                    await LoadAuthorizedOrdersAsync();
                    await LoadOrderDetailsAsync(OrderHeader.OrderNo);
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Save Failed",
                        Detail = "An error occurred while saving the PPC plan. Please try again.",
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }
    }
}
