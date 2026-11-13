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
        public string OrderFilterMode { get; set; } = "All"; // All | ActionNeeded | Planned
        public PpcOrderSummaryListItem? SelectedOrderItem { get; set; }
        public string SelectedOrderNo { get; set; } = string.Empty;
        public PpcOrderHeaderDto? OrderHeader { get; set; }
        public string PlanningNotes { get; set; } = string.Empty;

        public string ActiveTab { get; set; } = "ToPlan"; // ToPlan | Planned
        public bool ShowAwaitingAuthSection { get; set; } = false;
        private readonly HashSet<string> _unlockedRevisionItemIds = new(StringComparer.OrdinalIgnoreCase);

        public bool IsLoading { get; set; } = false;
        public bool IsSaving { get; set; } = false;

        public IEnumerable<PpcOrderSummaryListItem> FilteredAuthorizedOrders => OrderFilterMode switch
        {
            "ActionNeeded" => AuthorizedOrders.Where(o => !o.IsPlanned),
            "Planned" => AuthorizedOrders.Where(o => o.IsPlanned),
            _ => AuthorizedOrders
        };

        public List<PpcOrderItemDto> ToPlanItems => OrderHeader?.Items
            .Where(i => i.Authorized && !i.IsLockedInUi)
            .ToList() ?? new();

        public List<PpcOrderItemDto> LockedPlannedItems => OrderHeader?.Items
            .Where(i => i.Authorized && i.IsLockedInUi)
            .ToList() ?? new();

        public List<PpcOrderItemDto> AwaitingAuthItems => OrderHeader?.Items
            .Where(i => !i.Authorized)
            .ToList() ?? new();

        public bool CanSavePlan =>
            OrderHeader != null &&
            OrderHeader.Items.Any(i => i.Authorized) &&
            OrderHeader.Items.Any(i => i.Authorized && (i.IsPlanned || i.HasDeltaQty)) &&
            !OrderHeader.Items.Where(i => i.Authorized).Any(i => i.StockQty > i.AvailableInHandStock || i.StockQty < i.MinAllowedStockQty) &&
            !OrderHeader.Items.Where(i => i.Authorized).Any(i => i.SFStockAllocations.Any(a => a.AllocatedQty < a.AlreadyIssuedQty || a.AllocatedQty > a.AvailableQty)) &&
            !OrderHeader.Items.Where(i => i.Authorized).Any(i => i.Purchases.Where(p => !string.Equals(p.Status, "Cancelled", StringComparison.OrdinalIgnoreCase)).Any(p => p.PurchaseQty <= 0 || p.ProcessID <= 0 || (p.IsIssued && p.PurchaseQty < p.ReceivedQty))) &&
            !OrderHeader.Items.Where(i => i.Authorized).Any(i => i.HubSchedules.Any(h => h.StartDate > h.EndDate));

        // ── SF Stock Modal State ───────────────────────────────────────────────
        public PpcOrderItemDto? ActiveSfItem { get; set; }
        public bool ShowSfModal { get; set; } = false;

        // ── Exception / Revision Modal State ───────────────────────────────────
        public PpcOrderItemDto? ExceptionItem { get; set; }
        public bool ShowExceptionModal { get; set; } = false;
        public string ExceptionType { get; set; } = "UnlockPlan";
        public int ExceptionAffectedQty { get; set; } = 0;
        public int? ExceptionSelectedPurchaseEntryId { get; set; }
        public int? ExceptionSelectedSfSooRefId { get; set; }
        public string ExceptionReason { get; set; } = string.Empty;
        public bool IsProcessingException { get; set; } = false;

        public void OpenSfModal(PpcOrderItemDto item)
        {
            if (!item.Authorized) return;

            foreach (var lot in item.AvailableSFLots)
            {
                if (!item.SFStockAllocations.Any(a => a.SOO_RefID == lot.EntryID))
                {
                    item.SFStockAllocations.Add(new PpcItemSFAllocationDto
                    {
                        SOO_RefID = lot.EntryID,
                        ProcessID = lot.ProcessID,
                        ProcessName = lot.ProcessName,
                        LotNo = lot.LotNo,
                        BatchNo = lot.BatchNo,
                        Shelf_RefID = lot.Shelf_RefID,
                        StoreLocation = lot.StoreLocation,
                        AvailableQty = lot.AvailableQty,
                        AllocatedQty = 0,
                        AlreadyIssuedQty = 0
                    });
                }
            }

            ActiveSfItem = item;
            ShowSfModal = true;
        }

        public void CloseSfModal()
        {
            if (ActiveSfItem != null)
            {
                ActiveSfItem.SFStockAllocations.RemoveAll(a => a.AllocatedQty <= 0 && a.AlreadyIssuedQty <= 0);
            }
            ShowSfModal = false;
            ActiveSfItem = null;
        }

        public void SetMaxSfAllocation(PpcItemSFAllocationDto alloc)
        {
            alloc.AllocatedQty = alloc.AvailableQty;
        }

        public void ClearSfAllocation(PpcItemSFAllocationDto alloc)
        {
            if (alloc.AlreadyIssuedQty > 0)
            {
                alloc.AllocatedQty = alloc.AlreadyIssuedQty;
                return;
            }
            alloc.AllocatedQty = 0;
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadAuthorizedOrdersAsync();

            string? initialOrderNo = !string.IsNullOrWhiteSpace(OrderNoParam)
                ? OrderNoParam
                : OrderNoQuery;

            if (!string.IsNullOrWhiteSpace(initialOrderNo))
            {
                SelectedOrderNo = initialOrderNo.Trim();
                SelectedOrderItem = AuthorizedOrders.FirstOrDefault(o => string.Equals(o.OrderNo, SelectedOrderNo, StringComparison.OrdinalIgnoreCase));
                await LoadOrderDetailsAsync(SelectedOrderNo);
            }
        }

        private async Task LoadAuthorizedOrdersAsync()
        {
            try
            {
                AuthorizedOrders = await PpcPlanningService.GetAuthorizedOrdersForPlanningAsync();
                if (!string.IsNullOrWhiteSpace(SelectedOrderNo))
                {
                    SelectedOrderItem = AuthorizedOrders.FirstOrDefault(o => string.Equals(o.OrderNo, SelectedOrderNo, StringComparison.OrdinalIgnoreCase));
                }
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

        public Task<IEnumerable<PpcOrderSummaryListItem>> SearchOrders(string searchText)
        {
            var baseList = FilteredAuthorizedOrders;
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult(baseList);

            string term = searchText.Trim();
            var filtered = baseList.Where(o =>
                (!string.IsNullOrEmpty(o.OrderNo) && o.OrderNo.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(o.CustomerName) && o.CustomerName.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(o.CustCode) && o.CustCode.Contains(term, StringComparison.OrdinalIgnoreCase)));

            return Task.FromResult(filtered);
        }

        public void SetOrderFilterMode(string mode)
        {
            OrderFilterMode = mode;
        }

        public async Task OnOrderSelectedAsync(PpcOrderSummaryListItem? order)
        {
            SelectedOrderItem = order;
            SelectedOrderNo = order?.OrderNo ?? string.Empty;
            _unlockedRevisionItemIds.Clear();

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
                await LoadAuthorizedOrdersAsync();
                await LoadOrderDetailsAsync(SelectedOrderNo);
            }
        }

        private async Task LoadOrderDetailsAsync(string orderNo, bool preserveActiveTab = false)
        {
            try
            {
                IsLoading = true;
                OrderHeader = await PpcPlanningService.GetOrderForPlanningAsync(orderNo);
                if (OrderHeader != null)
                {
                    PlanningNotes = OrderHeader.Notes ?? string.Empty;

                    foreach (var item in OrderHeader.Items)
                    {
                        if (_unlockedRevisionItemIds.Contains(item.CompItemID))
                        {
                            item.IsUnlockedForRevision = true;
                        }

                        foreach (var hub in item.HubSchedules)
                        {
                            if (hub.PlannedQty == 0)
                            {
                                hub.PlannedQty = item.ProductionQty;
                            }
                        }
                    }

                    if (!preserveActiveTab)
                    {
                        ActiveTab = OrderHeader.ActionRequiredItemsCount > 0 ? "ToPlan" : "Planned";
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

        public async Task LockItemRevisionAsync(PpcOrderItemDto item)
        {
            _unlockedRevisionItemIds.Remove(item.CompItemID);
            item.IsUnlockedForRevision = false;
            if (!string.IsNullOrWhiteSpace(SelectedOrderNo))
            {
                await LoadOrderDetailsAsync(SelectedOrderNo, preserveActiveTab: false);
            }
        }

        // ── Exception Modal Logic ──────────────────────────────────────────────
        public void OpenExceptionModal(PpcOrderItemDto item, string? presetExceptionType = null, int? presetPurchaseEntryId = null)
        {
            if (!item.Authorized) return;

            ExceptionItem = item;
            ExceptionReason = string.Empty;
            ExceptionSelectedPurchaseEntryId = presetPurchaseEntryId;
            ExceptionSelectedSfSooRefId = item.SFStockAllocations.FirstOrDefault(a => a.AllocatedQty > a.AlreadyIssuedQty)?.SOO_RefID;

            if (!string.IsNullOrEmpty(presetExceptionType))
            {
                ExceptionType = presetExceptionType;
            }
            else if (item.Purchases.Any(p => p.CanCancelIssuedPo))
            {
                ExceptionType = "CancelMakerPO";
                ExceptionSelectedPurchaseEntryId ??= item.Purchases.First(p => p.CanCancelIssuedPo).EntryID;
            }
            else if (item.Purchases.Any(p => p.CanShortCloseIssuedPo))
            {
                ExceptionType = "ShortCloseMakerPO";
                ExceptionSelectedPurchaseEntryId ??= item.Purchases.First(p => p.CanShortCloseIssuedPo).EntryID;
            }
            else
            {
                ExceptionType = "UnlockPlan";
            }

            UpdateDefaultExceptionQty();
            ShowExceptionModal = true;
        }

        public void OnExceptionTypeChanged(string newType)
        {
            ExceptionType = newType;
            if (ExceptionItem != null)
            {
                if (newType == "CancelMakerPO")
                {
                    ExceptionSelectedPurchaseEntryId = ExceptionItem.Purchases.FirstOrDefault(p => p.CanCancelIssuedPo || !p.IsIssued)?.EntryID;
                }
                else if (newType == "ShortCloseMakerPO")
                {
                    ExceptionSelectedPurchaseEntryId = ExceptionItem.Purchases.FirstOrDefault(p => p.CanShortCloseIssuedPo)?.EntryID;
                }
                else if (newType == "SFStockRelease")
                {
                    ExceptionSelectedSfSooRefId = ExceptionItem.SFStockAllocations.FirstOrDefault(a => a.AllocatedQty > a.AlreadyIssuedQty)?.SOO_RefID;
                }
            }
            UpdateDefaultExceptionQty();
        }

        public void OnExceptionSfLotChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int sooRefId))
            {
                ExceptionSelectedSfSooRefId = sooRefId;
                UpdateDefaultExceptionQty();
            }
        }

        private void UpdateDefaultExceptionQty()
        {
            if (ExceptionItem == null)
            {
                ExceptionAffectedQty = 0;
                return;
            }

            switch (ExceptionType)
            {
                case "FinishStockRelease":
                    ExceptionAffectedQty = Math.Max(0, ExceptionItem.StockQty - ExceptionItem.AlreadyAdjustedStockQty);
                    break;
                case "FinishStockRejection":
                    ExceptionAffectedQty = Math.Max(0, ExceptionItem.AlreadyAdjustedStockQty - ExceptionItem.RejectedStockQty);
                    break;
                case "SFStockRelease":
                {
                    var alloc = ExceptionItem.SFStockAllocations.FirstOrDefault(a => a.SOO_RefID == ExceptionSelectedSfSooRefId);
                    ExceptionAffectedQty = alloc != null ? Math.Max(0, alloc.AllocatedQty - alloc.AlreadyIssuedQty) : 0;
                    break;
                }
                case "SFStockRejection":
                    ExceptionAffectedQty = Math.Max(0, ExceptionItem.SFStockQty - ExceptionItem.RejectedSFQty);
                    break;
                case "CancelMakerPO":
                {
                    var p = ExceptionItem.Purchases.FirstOrDefault(x => x.EntryID == ExceptionSelectedPurchaseEntryId);
                    ExceptionAffectedQty = p?.PurchaseQty ?? 0;
                    break;
                }
                case "ShortCloseMakerPO":
                {
                    var p = ExceptionItem.Purchases.FirstOrDefault(x => x.EntryID == ExceptionSelectedPurchaseEntryId);
                    ExceptionAffectedQty = p != null ? Math.Max(0, p.PurchaseQty - p.ReceivedQty) : 0;
                    break;
                }
                default:
                    ExceptionAffectedQty = 0;
                    break;
            }
        }

        public void CloseExceptionModal()
        {
            ShowExceptionModal = false;
            ExceptionItem = null;
            ExceptionReason = string.Empty;
        }

        public async Task SubmitExceptionAsync()
        {
            if (OrderHeader == null || ExceptionItem == null || IsProcessingException) return;

            if (string.IsNullOrWhiteSpace(ExceptionReason))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Reason Required",
                    Detail = "Please provide a clear reason for this exception/revision.",
                    Duration = 4000
                });
                return;
            }

            IsProcessingException = true;
            try
            {
                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                string currentUserName = authState.User.Identity?.Name ?? "PPC Planner";

                var request = new PpcExceptionActionRequest
                {
                    OrderNo = OrderHeader.OrderNo,
                    ItemID = ExceptionItem.CompItemID,
                    ExceptionType = ExceptionType,
                    AffectedQty = ExceptionAffectedQty,
                    PurchaseSplitEntryID = ExceptionSelectedPurchaseEntryId,
                    SfAllocationSooRefID = ExceptionSelectedSfSooRefId,
                    Reason = ExceptionReason.Trim(),
                    ActionBy = currentUserName
                };

                var result = await PpcPlanningService.ExecutePlanningExceptionAsync(request);
                if (result.Success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Exception Applied",
                        Detail = result.Message,
                        Duration = 5000
                    });

                    string targetItemId = ExceptionItem.CompItemID;
                    _unlockedRevisionItemIds.Add(targetItemId);
                    CloseExceptionModal();

                    await LoadAuthorizedOrdersAsync();
                    await LoadOrderDetailsAsync(OrderHeader.OrderNo, preserveActiveTab: true);
                    ActiveTab = "ToPlan";
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Exception Failed",
                        Detail = result.Message,
                        Duration = 5000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsProcessingException = false;
            }
        }

        // ── Purchase Splits ────────────────────────────────────────────────────
        public async Task AddPurchaseSplit(PpcOrderItemDto item)
        {
            if (!item.Authorized || item.IsLockedInUi) return;

            var split = new PpcItemPurchaseAllocationDto
            {
                ProcessID = 0,
                ProcessName = "",
                VendID = null,
                MakerName = null,
                PurchaseQty = Math.Max(0, item.ProductionQty),
                PurchaseRate = 0,
                Status = "Active"
            };

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
            if (split.IsIssued) return;
            item.Purchases.Remove(split);
        }

        public async Task OnPurchaseProcessChanged(PpcOrderItemDto item, PpcItemPurchaseAllocationDto split, ChangeEventArgs e)
        {
            if (split.IsIssued) return;
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
            if (split.IsIssued) return;
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
            if (IsSaving || OrderHeader == null) return;

            var authorizedItems = OrderHeader.Items.Where(i => i.Authorized).ToList();
            if (!authorizedItems.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Authorized Items",
                    Detail = "This order has no authorized items to plan. Material planning is restricted until authorized by a Director.",
                    Duration = 5000
                });
                return;
            }

            var plannedItems = authorizedItems.Where(i => i.IsPlanned || i.HasDeltaQty).ToList();
            if (!plannedItems.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Items Planned",
                    Detail = "Please plan at least one item (add a maker purchase split, issue stock, or mark as in-house planned) before saving.",
                    Duration = 5000
                });
                return;
            }

            foreach (var item in plannedItems)
            {
                if (item.StockQty < item.MinAllowedStockQty)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Finish Stock Floor Violated",
                        Detail = $"Finish Stock for {item.CompItemID} cannot be reduced below {item.MinAllowedStockQty} pcs (already withdrawn/rejected in store). Use 'Handle Exception' to record QC rejection.",
                        Duration = 5000
                    });
                    return;
                }

                if (item.StockQty > item.AvailableInHandStock)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Invalid Finish Stock Qty",
                        Detail = $"Finish Stock quantity for {item.CompItemID} ({item.StockQty}) exceeds available finish stock ({item.AvailableInHandStock}).",
                        Duration = 5000
                    });
                    return;
                }

                foreach (var sfAlloc in item.SFStockAllocations)
                {
                    if (sfAlloc.AllocatedQty < sfAlloc.AlreadyIssuedQty)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Warning,
                            Summary = "SF Allocation Floor Violated",
                            Detail = $"Lot {sfAlloc.LotNo} for {item.CompItemID} cannot be reduced below already-issued quantity ({sfAlloc.AlreadyIssuedQty} pcs).",
                            Duration = 5000
                        });
                        return;
                    }

                    if (sfAlloc.AllocatedQty > sfAlloc.AvailableQty)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Warning,
                            Summary = "Invalid SF Allocation",
                            Detail = $"Allocated quantity ({sfAlloc.AllocatedQty}) exceeds available SF stock ({sfAlloc.AvailableQty}) for lot {sfAlloc.LotNo} on item {item.CompItemID}.",
                            Duration = 5000
                        });
                        return;
                    }
                }

                var activePurchases = item.Purchases.Where(p => !string.Equals(p.Status, "Cancelled", StringComparison.OrdinalIgnoreCase)).ToList();
                if (activePurchases.Any(p => p.PurchaseQty <= 0 || p.ProcessID <= 0))
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

                foreach (var p in activePurchases.Where(x => x.IsIssued))
                {
                    if (p.PurchaseQty < p.ReceivedQty)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Warning,
                            Summary = "PO Qty Floor Violated",
                            Detail = $"PO #{p.MasterPONo} for {item.CompItemID} cannot be less than already received qty ({p.ReceivedQty} pcs).",
                            Duration = 5000
                        });
                        return;
                    }
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
                    Items = authorizedItems
                };

                bool ok = await PpcPlanningService.SavePpcPlanAndDispatchTasksAsync(request);

                if (ok)
                {
                    _unlockedRevisionItemIds.Clear();

                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "PPC Plan Saved & Dispatched",
                        Detail = $"Plan for Order #{OrderHeader.OrderNo} saved successfully. Tasks dispatched to Stock, Purchaser, and Hub Supervisors.",
                        Duration = 5000
                    });

                    await LoadAuthorizedOrdersAsync();
                    await LoadOrderDetailsAsync(OrderHeader.OrderNo, preserveActiveTab: false);
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
