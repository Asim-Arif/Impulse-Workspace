using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
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
    }
}
