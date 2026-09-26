using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Http;
using Radzen;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.Stock;

namespace Impulse.Pages.Stock.StockOrderAdjustment
{
    public partial class StockOrderAdjustment : ComponentBase
    {
        [Inject] public IPpcStockAdjustmentService PpcStockAdjustmentService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] public AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;
        [Inject] public IHttpContextAccessor HttpContextAccessor { get; set; } = default!;

        [Parameter] public string? OrderNo { get; set; }
        [SupplyParameterFromQuery(Name = "orderNo")] public string? QueryOrderNo { get; set; }
        [SupplyParameterFromQuery(Name = "returnUrl")] public string? ReturnUrl { get; set; }

        public string? EffectiveOrderNo => !string.IsNullOrWhiteSpace(OrderNo) ? OrderNo : QueryOrderNo;

        public bool IsLoading { get; set; } = true;
        public bool IsSaving { get; set; } = false;

        public DateTime AdjustmentDate { get; set; } = DateTime.Today;

        public PpcStockAdjustmentOrderHeaderModel? OrderHeader { get; set; }
        public List<PpcStockAdjustmentPlannedItemModel> PlannedItems { get; set; } = new();

        public List<PpcStockAdjustmentOrderHeaderModel> ActiveOrdersWithStock { get; set; } = new();
        public PpcStockAdjustmentOrderHeaderModel? SelectedOrderLookup { get; set; }

        public string CurrentUserName { get; set; } = "InventorySupervisor";
        public string CurrentMachineName { get; set; } = "WEB-CLIENT";

        // Modal States
        public bool IsShelfModalOpen { get; set; } = false;
        public PpcStockAdjustmentPlannedItemModel? ModalItem { get; set; }
        public int ModalItemAllocatedSum => ModalItem?.Allocations?.Sum(a => a.AllocatedQty) ?? 0;

        public bool IsResultModalOpen { get; set; } = false;
        public PpcStockAdjustmentResultModel? LastResult { get; set; }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoading = true;

                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                var user = authState.User;
                if (user.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = user.Identity.Name ?? "InventorySupervisor";
                }

                CurrentMachineName = HttpContextAccessor.HttpContext?.Connection?.RemoteIpAddress?.ToString() ?? Environment.MachineName;

                ActiveOrdersWithStock = await PpcStockAdjustmentService.GetActiveOrdersWithStockPlanningAsync();

                if (!string.IsNullOrWhiteSpace(EffectiveOrderNo))
                {
                    await LoadOrderDataAsync(EffectiveOrderNo);
                }
                else if (ActiveOrdersWithStock.Any())
                {
                    var firstPending = ActiveOrdersWithStock.FirstOrDefault(o => o.TotalRemainingStockQty > 0)
                                      ?? ActiveOrdersWithStock.First();
                    await LoadOrderDataAsync(firstPending.OrderNo);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Initialization Error",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected override async Task OnParametersSetAsync()
        {
            if (!string.IsNullOrWhiteSpace(EffectiveOrderNo) &&
                (OrderHeader == null || !string.Equals(OrderHeader.OrderNo, EffectiveOrderNo, StringComparison.OrdinalIgnoreCase)))
            {
                await LoadOrderDataAsync(EffectiveOrderNo);
            }
        }

        public async Task LoadOrderDataAsync(string orderNo)
        {
            try
            {
                IsLoading = true;
                StateHasChanged();

                OrderHeader = await PpcStockAdjustmentService.GetOrderHeaderAsync(orderNo);
                SelectedOrderLookup = ActiveOrdersWithStock.FirstOrDefault(o => o.OrderNo == orderNo) ?? OrderHeader;

                if (OrderHeader != null)
                {
                    PlannedItems = await PpcStockAdjustmentService.GetPlannedStockItemsAsync(orderNo);
                }
                else
                {
                    PlannedItems = new List<PpcStockAdjustmentPlannedItemModel>();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Order",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        public async Task<IEnumerable<PpcStockAdjustmentOrderHeaderModel>> SearchOrdersAsync(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(ActiveOrdersWithStock.Take(25));
            }

            var query = ActiveOrdersWithStock.Where(o =>
                o.OrderNo.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                o.CustomerName.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                o.CustPONo.Contains(searchText, StringComparison.OrdinalIgnoreCase));

            return await Task.FromResult(query.Take(25));
        }

        public async Task OnOrderSelected(PpcStockAdjustmentOrderHeaderModel? selected)
        {
            SelectedOrderLookup = selected;
            if (selected != null)
            {
                NavigationManager.NavigateTo($"/stock/stock-order-adjustment?orderNo={selected.OrderNo}");
                await LoadOrderDataAsync(selected.OrderNo);
            }
        }

        public async Task AutoAllocateAllStock()
        {
            foreach (var item in PlannedItems.Where(i => i.RemainingQtyToIssue > 0))
            {
                await PpcStockAdjustmentService.AutoAllocateStockAsync(item);
            }

            NotificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Info,
                Summary = "Auto-Allocation Complete",
                Detail = "Available in-hand stock has been allocated FIFO across items.",
                Duration = 3000
            });
            StateHasChanged();
        }

        public async Task OnAdjQtyChanged(PpcStockAdjustmentPlannedItemModel item, ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int newQty))
            {
                if (newQty < 0) newQty = 0;
                if (newQty > item.RemainingQtyToIssue)
                {
                    newQty = item.RemainingQtyToIssue;
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Exceeded Planned Qty",
                        Detail = $"Issuance cannot exceed remaining planned qty of {item.RemainingQtyToIssue}.",
                        Duration = 4000
                    });
                }

                item.AdjQty = newQty;
                await PpcStockAdjustmentService.AutoAllocateStockAsync(item);
            }
        }

        // Modal Allocation Management
        public void OpenShelfAllocationModal(PpcStockAdjustmentPlannedItemModel item)
        {
            ModalItem = item;
            if (ModalItem.Allocations == null) ModalItem.Allocations = new List<PpcStockShelfAllocationModel>();
            IsShelfModalOpen = true;
        }

        public void CloseShelfAllocationModal()
        {
            if (ModalItem != null && ModalItemAllocatedSum != ModalItem.AdjQty)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Allocation Mismatch",
                    Detail = $"Allocated quantity ({ModalItemAllocatedSum}) does not match Issue Qty ({ModalItem.AdjQty}).",
                    Duration = 4000
                });
            }
            IsShelfModalOpen = false;
            ModalItem = null;
        }

        public void OnShelfAllocQtyChanged(PpcStockShelfAllocationModel lot, ChangeEventArgs e)
        {
            if (ModalItem == null) return;

            if (int.TryParse(e.Value?.ToString(), out int enteredQty))
            {
                if (enteredQty < 0) enteredQty = 0;
                if (enteredQty > (int)lot.NetAvailableQty) enteredQty = (int)lot.NetAvailableQty;

                var existing = ModalItem.Allocations.FirstOrDefault(a => a.Shelf_RefID == lot.Shelf_RefID && a.LotNo == lot.LotNo);
                if (existing != null)
                {
                    existing.AllocatedQty = enteredQty;
                }
                else if (enteredQty > 0)
                {
                    ModalItem.Allocations.Add(new PpcStockShelfAllocationModel
                    {
                        Shelf_RefID = lot.Shelf_RefID,
                        StoreName = lot.StoreName,
                        RackNo = lot.RackNo,
                        ShelfNo = lot.ShelfNo,
                        LotNo = lot.LotNo,
                        Batch_No = lot.Batch_No,
                        Mill_Certificate_No = lot.Mill_Certificate_No,
                        RCV_ISD_RefID = lot.RCV_ISD_RefID,
                        NetAvailableQty = lot.NetAvailableQty,
                        AllocatedQty = enteredQty
                    });
                }
            }
        }

        public void SetMaxShelfAlloc(PpcStockShelfAllocationModel lot)
        {
            if (ModalItem == null) return;

            int currentOthers = ModalItem.Allocations
                .Where(a => !(a.Shelf_RefID == lot.Shelf_RefID && a.LotNo == lot.LotNo))
                .Sum(a => a.AllocatedQty);

            int remainingNeeded = Math.Max(0, ModalItem.AdjQty - currentOthers);
            int take = Math.Min(remainingNeeded, (int)lot.NetAvailableQty);

            var existing = ModalItem.Allocations.FirstOrDefault(a => a.Shelf_RefID == lot.Shelf_RefID && a.LotNo == lot.LotNo);
            if (existing != null)
            {
                existing.AllocatedQty = take;
            }
            else
            {
                ModalItem.Allocations.Add(new PpcStockShelfAllocationModel
                {
                    Shelf_RefID = lot.Shelf_RefID,
                    StoreName = lot.StoreName,
                    RackNo = lot.RackNo,
                    ShelfNo = lot.ShelfNo,
                    LotNo = lot.LotNo,
                    Batch_No = lot.Batch_No,
                    Mill_Certificate_No = lot.Mill_Certificate_No,
                    RCV_ISD_RefID = lot.RCV_ISD_RefID,
                    NetAvailableQty = lot.NetAvailableQty,
                    AllocatedQty = take
                });
            }
        }

        public async Task AutoAllocateCurrentModalItem()
        {
            if (ModalItem != null)
            {
                await PpcStockAdjustmentService.AutoAllocateStockAsync(ModalItem);
            }
        }

        public void ClearCurrentModalAllocations()
        {
            if (ModalItem != null)
            {
                ModalItem.Allocations.Clear();
            }
        }

        public async Task SaveStockAdjustmentAsync()
        {
            if (OrderHeader == null || !PlannedItems.Any()) return;

            var itemsToSave = PlannedItems.Where(i => i.AdjQty > 0).ToList();
            if (!itemsToSave.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Quantities Entered",
                    Detail = "Please enter an issuance quantity for at least one planned item.",
                    Duration = 4000
                });
                return;
            }

            // Validation check
            foreach (var item in itemsToSave)
            {
                if (!item.IsAllocationValid)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Validation Failed",
                        Detail = $"Item {item.ItemID} has an allocation mismatch ({item.TotalAllocatedQty} vs {item.AdjQty}) or missing target process.",
                        Duration = 6000
                    });
                    return;
                }
            }

            try
            {
                IsSaving = true;

                var submission = new PpcStockAdjustmentSubmissionModel
                {
                    OrderNo = OrderHeader.OrderNo,
                    AdjustmentDate = AdjustmentDate,
                    UserName = CurrentUserName,
                    MachineName = CurrentMachineName,
                    Items = itemsToSave
                };

                var result = await PpcStockAdjustmentService.SaveStockAdjustmentAsync(submission);

                if (result.Success)
                {
                    LastResult = result;
                    IsResultModalOpen = true;

                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Stock Issued Successfully",
                        Detail = $"Stock successfully adjusted for Order #{OrderHeader.OrderNo}.",
                        Duration = 5000
                    });

                    // Refresh active orders list and current order data
                    ActiveOrdersWithStock = await PpcStockAdjustmentService.GetActiveOrdersWithStockPlanningAsync();
                    await LoadOrderDataAsync(OrderHeader.OrderNo);
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Save Failed",
                        Detail = result.ErrorMessage,
                        Duration = 8000
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
                    Duration = 8000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public void CloseResultModal()
        {
            IsResultModalOpen = false;
            LastResult = null;
        }

        public void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/stock/finish-stock-issuance");
            }
        }
    }
}
