using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Components.Production.MakerBilling
{
    public partial class HiddenLotsTab : ComponentBase
    {
        [Inject] private IMakerBillingService BillingService { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public List<MakerHiddenLotItem> Items { get; set; } = new List<MakerHiddenLotItem>();
        [Parameter] public bool IsLoading { get; set; } = false;
        [Parameter] public EventCallback OnLotRestored { get; set; }
        [Parameter] public EventCallback<MakerHiddenLotItem> OnPrintPTC { get; set; }

        public MakerHiddenLotItem? ContextRowItem { get; set; } = null;

        public async Task RestoreLot(MakerHiddenLotItem item)
        {
            try
            {
                await BillingService.RestoreHiddenLotAsync(item.EntryID);
                Items.Remove(item);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Lot Restored",
                    Detail = $"Lot #{item.LotNo} restored to Maker Billing.",
                    Duration = 3000
                });

                await OnLotRestored.InvokeAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Restore Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task ContextRestoreLot(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await RestoreLot(ContextRowItem);
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
            if (args.Data is MakerHiddenLotItem item)
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
