using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Components.Production.MakerBilling
{
    public partial class MakerBillingTab : ComponentBase
    {
        [Inject] private IMakerBillingService BillingService { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public List<MakerBillingLotItem> Lots { get; set; } = new List<MakerBillingLotItem>();
        [Parameter] public bool IsLoading { get; set; } = false;
        [Parameter] public bool ShowPicture { get; set; } = false;
        [Parameter] public bool RejectionRelaxation { get; set; } = false;

        [Parameter] public decimal ShortTermDeduction { get; set; } = 0;
        [Parameter] public EventCallback<decimal> ShortTermDeductionChanged { get; set; }

        [Parameter] public decimal LongTermDeduction { get; set; } = 0;
        [Parameter] public EventCallback<decimal> LongTermDeductionChanged { get; set; }

        [Parameter] public decimal RepairAmount { get; set; } = 0;
        [Parameter] public decimal MaterialDeductionAmount { get; set; } = 0;
        [Parameter] public decimal ManualDeductionAmount { get; set; } = 0;

        [Parameter] public EventCallback OnStateUpdated { get; set; }
        [Parameter] public EventCallback<MakerBillingLotItem> OnHideLot { get; set; }
        [Parameter] public EventCallback<MakerBillingLotItem> OnPrintPTC { get; set; }

        // Local Filter & State
        public string SearchLotText { get; set; } = string.Empty;
        public bool IsAllSelected { get; set; } = false;
        public long EditingLotId { get; set; } = 0;
        public decimal EditingRate { get; set; } = 0;
        public string HoveredImage { get; set; } = string.Empty;
        public MakerBillingLotItem? ContextRowItem { get; set; } = null;

        // Computed Properties
        public IEnumerable<MakerBillingLotItem> FilteredLots => string.IsNullOrWhiteSpace(SearchLotText)
            ? Lots
            : Lots.Where(x => x.LotNo.Contains(SearchLotText, StringComparison.OrdinalIgnoreCase) ||
                              x.ItemCode.Contains(SearchLotText, StringComparison.OrdinalIgnoreCase) ||
                              x.ItemName.Contains(SearchLotText, StringComparison.OrdinalIgnoreCase));

        public List<MakerBillingLotItem> SelectedLots => Lots.Where(x => x.IsSelected).ToList();
        public int AuditedLotsCount => Lots.Count(x => x.Audited);
        public decimal TotalSelectedQty => SelectedLots.Sum(x => x.Qty);
        public decimal TotalSelectedValue => SelectedLots.Sum(x => x.Value);

        public decimal RejectionRelaxAmount => RejectionRelaxation ? (TotalSelectedValue / 100m) * 2m : 0;
        public decimal EffectiveRepairDeduction => Math.Max(0, RepairAmount - RejectionRelaxAmount);

        public decimal NetPayableAmount => TotalSelectedValue
            - ShortTermDeduction
            - LongTermDeduction
            - MaterialDeductionAmount
            - ManualDeductionAmount
            - EffectiveRepairDeduction;

        public void SelectAll()
        {
            foreach (var item in Lots) item.IsSelected = true;
            IsAllSelected = true;
            OnStateUpdated.InvokeAsync();
        }

        public void UnselectAll()
        {
            foreach (var item in Lots) item.IsSelected = false;
            IsAllSelected = false;
            OnStateUpdated.InvokeAsync();
        }

        public void SelectAudited()
        {
            foreach (var item in Lots) item.IsSelected = item.Audited;
            IsAllSelected = Lots.Any() && Lots.All(x => x.IsSelected);
            OnStateUpdated.InvokeAsync();
        }

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            IsAllSelected = (bool)(e.Value ?? false);
            foreach (var item in Lots) item.IsSelected = IsAllSelected;
            OnStateUpdated.InvokeAsync();
        }

        public void OnSelectionChanged()
        {
            IsAllSelected = Lots.Any() && Lots.All(x => x.IsSelected);
            OnStateUpdated.InvokeAsync();
        }

        public async Task MarkAsAudited()
        {
            var selectedIds = SelectedLots.Select(x => x.EntryID).ToList();
            if (!selectedIds.Any()) return;

            try
            {
                await BillingService.MarkLotsAsAuditedAsync(selectedIds, true);
                foreach (var lot in SelectedLots) lot.Audited = true;

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Marked Audited",
                    Detail = $"{selectedIds.Count} lot(s) marked as audited.",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Update Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task UnmarkAudited()
        {
            var selectedIds = SelectedLots.Select(x => x.EntryID).ToList();
            if (!selectedIds.Any()) return;

            try
            {
                await BillingService.MarkLotsAsAuditedAsync(selectedIds, false);
                foreach (var lot in SelectedLots) lot.Audited = false;

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Info,
                    Summary = "Unmarked Audited",
                    Detail = $"{selectedIds.Count} lot(s) unmarked.",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Update Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public void StartRateEdit(MakerBillingLotItem lot)
        {
            EditingLotId = lot.EntryID;
            EditingRate = lot.Rate;
        }

        public void CancelRateEdit()
        {
            EditingLotId = 0;
        }

        public async Task SaveRate(MakerBillingLotItem lot)
        {
            if (EditingRate <= 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Invalid Rate",
                    Detail = "Rate must be greater than zero.",
                    Duration = 3000
                });
                return;
            }

            try
            {
                await BillingService.UpdateLotRateAsync(lot.VI_EntryID, EditingRate);
                lot.Rate = EditingRate;

                // Update Rcvd. Value (BillingProcAmount in legacy is based on RcvdQty * Rate)
                decimal effectiveQty = lot.RcvdQty != 0 ? lot.RcvdQty : lot.IssQty;
                lot.RcvdValue = Math.Round(effectiveQty * EditingRate, 2);
                lot.BalValue = Math.Round(lot.RcvdValue - lot.UnAuthValue - lot.RepairAmt, 2);

                EditingLotId = 0;

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Rate Updated",
                    Detail = $"Rate for Lot #{lot.LotNo} updated to {EditingRate:N2}, Rcvd Value recalculated to {lot.RcvdValue:N2}.",
                    Duration = 3000
                });

                await OnStateUpdated.InvokeAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Rate Update Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task OnRateKeyDown(KeyboardEventArgs e, MakerBillingLotItem lot)
        {
            if (e.Key == "Enter")
            {
                await SaveRate(lot);
            }
            else if (e.Key == "Escape")
            {
                CancelRateEdit();
            }
        }

        public async Task OnDeductionChanged()
        {
            await ShortTermDeductionChanged.InvokeAsync(ShortTermDeduction);
            await LongTermDeductionChanged.InvokeAsync(LongTermDeduction);
            await OnStateUpdated.InvokeAsync();
        }

        public async Task ShowImagePreview(string? itemCode)
        {
            if (ShowPicture && !string.IsNullOrWhiteSpace(itemCode))
            {
                HoveredImage = await BillingService.GetItemPicBase64Async(itemCode) ?? string.Empty;
            }
        }

        public void HideImagePreview()
        {
            // Do not clear immediately on mouseleave so user can view while selecting
        }

        public void ContextEditRate(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                StartRateEdit(ContextRowItem);
            }
        }

        public async Task ContextHideLot(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await OnHideLot.InvokeAsync(ContextRowItem);
            }
        }

        public async Task ContextPrintPTC(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await OnPrintPTC.InvokeAsync(ContextRowItem);
            }
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is MakerBillingLotItem item)
            {
                ContextRowItem = item;
            }
            else
            {
                ContextRowItem = null;
            }
        }
    }
}
