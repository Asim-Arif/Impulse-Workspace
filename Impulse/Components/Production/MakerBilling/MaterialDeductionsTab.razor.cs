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
    public partial class MaterialDeductionsTab : ComponentBase
    {
        [Inject] private IMakerBillingService BillingService { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public List<MakerMaterialDedItem> Items { get; set; } = new List<MakerMaterialDedItem>();
        [Parameter] public bool IsLoading { get; set; } = false;
        [Parameter] public EventCallback<decimal> OnMaterialAmountChanged { get; set; }
        [Parameter] public EventCallback<MakerMaterialDedItem> OnHideRM { get; set; }
        [Parameter] public EventCallback OnPrintReport { get; set; }

        public bool IsAllSelected { get; set; } = true;
        public long EditingItemId { get; set; } = 0;
        public decimal EditingPrice { get; set; } = 0;
        public MakerMaterialDedItem? ContextRowItem { get; set; } = null;

        public decimal TotalSelectedMaterialAmount => Items.Where(x => x.IsSelected).Sum(x => x.Amount);

        public void SelectAll()
        {
            foreach (var item in Items) item.IsSelected = true;
            IsAllSelected = true;
            OnMaterialAmountChanged.InvokeAsync(TotalSelectedMaterialAmount);
        }

        public void UnselectAll()
        {
            foreach (var item in Items) item.IsSelected = false;
            IsAllSelected = false;
            OnMaterialAmountChanged.InvokeAsync(TotalSelectedMaterialAmount);
        }

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            IsAllSelected = (bool)(e.Value ?? false);
            foreach (var item in Items) item.IsSelected = IsAllSelected;
            OnMaterialAmountChanged.InvokeAsync(TotalSelectedMaterialAmount);
        }

        public void OnSelectionChanged()
        {
            IsAllSelected = Items.Any() && Items.All(x => x.IsSelected);
            OnMaterialAmountChanged.InvokeAsync(TotalSelectedMaterialAmount);
        }

        public void StartPriceEdit(MakerMaterialDedItem item)
        {
            EditingItemId = item.EntryID;
            EditingPrice = item.Price;
        }

        public void CancelPriceEdit()
        {
            EditingItemId = 0;
        }

        public async Task SavePrice(MakerMaterialDedItem item)
        {
            try
            {
                await BillingService.UpdateMaterialPriceAsync(item.EntryID, EditingPrice);
                item.Price = EditingPrice;
                EditingItemId = 0;

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Price Updated",
                    Detail = $"Price for {item.RawMaterialDisplay} updated to {EditingPrice:N2}.",
                    Duration = 3000
                });

                await OnMaterialAmountChanged.InvokeAsync(TotalSelectedMaterialAmount);
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

        public async Task OnPriceKeyDown(KeyboardEventArgs e, MakerMaterialDedItem item)
        {
            if (e.Key == "Enter")
            {
                await SavePrice(item);
            }
            else if (e.Key == "Escape")
            {
                CancelPriceEdit();
            }
        }

        public void ContextEditPrice(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                StartPriceEdit(ContextRowItem);
            }
        }

        public async Task ContextHideRM(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await OnHideRM.InvokeAsync(ContextRowItem);
            }
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is MakerMaterialDedItem item)
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
