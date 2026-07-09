using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;

namespace Impulse.Components.Export.NewProformaInvoiceTabs
{
    public partial class DetailsTab
    {
        [Parameter] public ProformaInvoiceViewModel Proforma { get; set; } = default!;
        [Parameter] public List<ProformaLineItemModel> LineItems { get; set; } = new();
        [Parameter] public EventCallback OnTotalsChanged { get; set; }
        [Parameter] public EventCallback<ProformaLineItemModel> OnRemoveItem { get; set; }

        // ── Inline edit state ──────────────────────────────────────────────
        private int EditingQtyIdx = -1;
        private int EditingPriceIdx = -1;
        private int EditingDescriptionIdx = -1;

        private double PriceUpdatePct;
        private string SearchTerm = string.Empty;

        // ── Context menu ───────────────────────────────────────────────────
        private bool ContextMenuVisible;
        private double ContextX;
        private double ContextY;
        private ProformaLineItemModel? ContextItem;

        // ── Computed ───────────────────────────────────────────────────────
        private int TotalPcs => LineItems.Sum(i => i.Qty);

        // ── Inline edit helpers ────────────────────────────────────────────

        private void StartEditQty(int idx) { CancelEdit(); EditingQtyIdx = idx; }
        private void StartEditPrice(int idx) { CancelEdit(); EditingPriceIdx = idx; }
        private void StartEditDescription(int idx) { CancelEdit(); EditingDescriptionIdx = idx; }

        private void CancelEdit()
        {
            EditingQtyIdx = -1;
            EditingPriceIdx = -1;
            EditingDescriptionIdx = -1;
        }

        private void CommitQty(int idx, string? val)
        {
            if (idx < 0 || idx >= LineItems.Count) return;
            if (int.TryParse(val, out int qty) && qty >= 0)
            {
                LineItems[idx].Qty = qty;
                RecalcTotals();
            }
            CancelEdit();
        }

        private void CommitPrice(int idx, string? val)
        {
            if (idx < 0 || idx >= LineItems.Count) return;
            if (double.TryParse(val, out double price) && price >= 0)
            {
                LineItems[idx].Price = price;
                RecalcTotals();
            }
            CancelEdit();
        }

        private void CommitDescription(int idx, string desc)
        {
            if (idx < 0 || idx >= LineItems.Count) return;
            LineItems[idx].Description = desc;
            CancelEdit();
        }

        private void HandleQtyKeyDown(KeyboardEventArgs e, int idx)
        {
            if (e.Key == "Escape") CancelEdit();
            // Enter handled by onchange
        }

        private void HandlePriceKeyDown(KeyboardEventArgs e, int idx)
        {
            if (e.Key == "Escape") CancelEdit();
        }

        private void HandleDescKeyDown(KeyboardEventArgs e, int idx)
        {
            if (e.Key == "Escape") CancelEdit();
        }

        // ── Totals recalculation ───────────────────────────────────────────

        private void RecalcTotals()
        {
            Proforma.TotalAmount = Math.Round(LineItems.Sum(i => i.Amount), 4);
            // Auto-update discount amount if percent is set
            if (Proforma.DiscountPer > 0)
                Proforma.DiscountAmount = Math.Round((Proforma.TotalAmount / 100.0) * Proforma.DiscountPer, 4);
            RecalcGrandTotal();
        }

        private void RecalcGrandTotal()
        {
            Proforma.GrandTotal = Math.Round(
                Proforma.TotalAmount
                - Proforma.DiscountAmount
                + Proforma.InsuranceFreightAmount
                + Proforma.InsuranceFreightAmountII, 4);
            OnTotalsChanged.InvokeAsync();
        }

        private void OnDiscountPerChanged(ChangeEventArgs e)
        {
            if (double.TryParse(e.Value?.ToString(), out double pct))
            {
                Proforma.DiscountPer = pct;
                Proforma.DiscountAmount = Math.Round((Proforma.TotalAmount / 100.0) * pct, 4);
                RecalcGrandTotal();
            }
        }

        private void OnDiscountAmtChanged(ChangeEventArgs e)
        {
            if (double.TryParse(e.Value?.ToString(), out double amt))
            {
                Proforma.DiscountAmount = amt;
                RecalcGrandTotal();
            }
        }

        // ── Price update ───────────────────────────────────────────────────

        private void ApplyPriceUpdate()
        {
            if (PriceUpdatePct == 0) return;
            for (int i = 0; i < LineItems.Count; i++)
            {
                double newPrice = LineItems[i].Price + (LineItems[i].Price / 100.0 * PriceUpdatePct);
                LineItems[i].Price = Math.Round(newPrice, 4);
            }
            RecalcTotals();
        }

        private async Task RemoveAllItems()
        {
            LineItems.Clear();
            RecalcTotals();
            await OnTotalsChanged.InvokeAsync();
        }

        private async Task RemoveItem(int idx)
        {
            if (idx < 0 || idx >= LineItems.Count) return;
            var item = LineItems[idx];
            await OnRemoveItem.InvokeAsync(item);
            LineItems.RemoveAt(idx);
            RecalcTotals();
        }

        // ── Context menu ───────────────────────────────────────────────────

        private void OnRowRightClick(MouseEventArgs e, ProformaLineItemModel item)
        {
            ContextItem = item;
            ContextX = e.ClientX;
            ContextY = e.ClientY;
            ContextMenuVisible = true;
        }

        private void HideContextMenu() => ContextMenuVisible = false;

        private void ToggleHighlight()
        {
            if (ContextItem == null) return;
            int idx = LineItems.IndexOf(ContextItem);
            if (idx >= 0)
                LineItems[idx].IsHighlighted = !LineItems[idx].IsHighlighted;
            HideContextMenu();
        }

        // ── Search ─────────────────────────────────────────────────────────

        private void OnSearchChanged(ChangeEventArgs e)
        {
            SearchTerm = e.Value?.ToString() ?? string.Empty;
            // In a real implementation, this would scroll/highlight the matching row
        }

        // Called from parent after items are added from NewItemsTab
        public void RefreshTotals() => RecalcTotals();
    }
}
