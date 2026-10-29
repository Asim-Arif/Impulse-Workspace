using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using BlazorContextMenu;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.VendorBilling.Tabs
{
    public partial class VendorBillingTab : ComponentBase
    {
        [Parameter] public List<VendorBillingRowViewModel> Rows { get; set; } = new();
        [Parameter] public bool IsLoading { get; set; }
        [Parameter] public EventCallback OnDataChanged { get; set; }
        [Parameter] public EventCallback<(int entryId, float newRate)> OnRateSaved { get; set; }
        [Parameter] public EventCallback<int> OnHideEntry { get; set; }

        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; }
        [Inject] private Microsoft.JSInterop.IJSRuntime JSRuntime { get; set; } = default!;

        public VendorBillingRowViewModel? ActiveRow { get; set; } = null;

        private VendorBillingRowViewModel? _contextMenuRow;

        private async Task OnRowRightClick(MouseEventArgs e, VendorBillingRowViewModel row)
        {
            _contextMenuRow = row;
            await BlazorContextMenuService.ShowMenu("billingRowContextMenu", (int)e.ClientX, (int)e.ClientY);
        }

        private async Task OnHideFromBilling(ItemClickEventArgs e)
        {
            if (_contextMenuRow != null)
            {
                await OnHideEntry.InvokeAsync(_contextMenuRow.EntryID);
                _contextMenuRow = null;
            }
        }

        private async Task OnRowChecked(VendorBillingRowViewModel row, ChangeEventArgs e)
        {
            row.IsChecked = (bool)(e.Value ?? false);
            await OnDataChanged.InvokeAsync();
        }

        private async Task OnSelectAllChanged(ChangeEventArgs e)
        {
            bool selectAll = (bool)(e.Value ?? false);
            foreach (var row in Rows)
                row.IsChecked = selectAll;
            await OnDataChanged.InvokeAsync();
        }

        private void StartEditRate(VendorBillingRowViewModel row)
        {
            row.EditingRate = row.Rate;
            row.IsEditingRate = true;
        }

        private void OnRateChanged(VendorBillingRowViewModel row, ChangeEventArgs e)
        {
            if (float.TryParse(e.Value?.ToString(), out float r))
                row.EditingRate = r;
        }

        private async Task OnRateKeyDown(VendorBillingRowViewModel row, KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
                await CommitRate(row);
            else if (e.Key == "Escape")
                row.IsEditingRate = false;
        }

        private async Task OnRateBlur(VendorBillingRowViewModel row)
        {
            await CommitRate(row);
        }

        private async Task CommitRate(VendorBillingRowViewModel row)
        {
            if (row.EditingRate > 0)
            {
                row.Rate = row.EditingRate;
                await OnRateSaved.InvokeAsync((row.EntryID, row.Rate));
            }
            row.IsEditingRate = false;
            await OnDataChanged.InvokeAsync();
        }

        protected override void OnParametersSet()
        {
            if (ActiveRow != null && !Rows.Contains(ActiveRow))
            {
                ActiveRow = Rows.FirstOrDefault();
            }
            else if (ActiveRow == null && Rows.Any())
            {
                ActiveRow = Rows.FirstOrDefault();
            }
        }

        public string GetRowClass(VendorBillingRowViewModel row)
        {
            var classes = new List<string>();
            if (row == ActiveRow)
            {
                classes.Add("table-row-active");
            }
            if (row.IsChecked)
            {
                classes.Add("table-primary");
            }
            return string.Join(" ", classes);
        }

        public void OnRowClicked(VendorBillingRowViewModel row)
        {
            ActiveRow = row;
        }

        public async Task HandleKeyDown(KeyboardEventArgs e)
        {
            if (Rows == null || !Rows.Any()) return;

            if (Rows.Any(r => r.IsEditingRate)) return;

            int currentIndex = ActiveRow != null ? Rows.IndexOf(ActiveRow) : -1;

            if (e.Key == "ArrowDown")
            {
                int newIndex = currentIndex < Rows.Count - 1 ? currentIndex + 1 : Rows.Count - 1;
                if (newIndex >= 0 && newIndex < Rows.Count)
                {
                    ActiveRow = Rows[newIndex];
                    await ScrollRowIntoView($"vendor-row-{ActiveRow.EntryID}");
                }
            }
            else if (e.Key == "ArrowUp")
            {
                int newIndex = currentIndex > 0 ? currentIndex - 1 : 0;
                if (newIndex >= 0 && newIndex < Rows.Count)
                {
                    ActiveRow = Rows[newIndex];
                    await ScrollRowIntoView($"vendor-row-{ActiveRow.EntryID}");
                }
            }
            else if (e.Key == " " || e.Code == "Space")
            {
                if (ActiveRow != null)
                {
                    ActiveRow.IsChecked = !ActiveRow.IsChecked;
                    await OnDataChanged.InvokeAsync();
                }
                else if (Rows.Any())
                {
                    ActiveRow = Rows[0];
                    ActiveRow.IsChecked = !ActiveRow.IsChecked;
                    await OnDataChanged.InvokeAsync();
                }
            }
            else if (e.Key == "Home")
            {
                ActiveRow = Rows[0];
                await ScrollRowIntoView($"vendor-row-{ActiveRow.EntryID}");
            }
            else if (e.Key == "End")
            {
                ActiveRow = Rows[Rows.Count - 1];
                await ScrollRowIntoView($"vendor-row-{ActiveRow.EntryID}");
            }
            else if (e.Key == "Enter")
            {
                if (ActiveRow != null && !ActiveRow.IsEditingRate)
                {
                    StartEditRate(ActiveRow);
                }
            }
        }

        private async Task ScrollRowIntoView(string rowId)
        {
            try
            {
                await JSRuntime.InvokeVoidAsync("scrollElementIntoView", rowId);
            }
            catch
            {
                // Ignored if DOM element not present
            }
        }
    }
}
