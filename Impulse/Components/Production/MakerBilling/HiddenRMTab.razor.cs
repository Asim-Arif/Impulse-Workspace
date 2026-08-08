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
    public partial class HiddenRMTab : ComponentBase
    {
        [Inject] private IMakerBillingService BillingService { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public List<MakerHiddenRMItem> Items { get; set; } = new List<MakerHiddenRMItem>();
        [Parameter] public bool IsLoading { get; set; } = false;
        [Parameter] public EventCallback OnRMRestored { get; set; }

        public MakerHiddenRMItem? ContextRowItem { get; set; } = null;

        public async Task RestoreRM(MakerHiddenRMItem item)
        {
            try
            {
                await BillingService.RestoreHiddenRMAsync(item.EntryID);
                Items.Remove(item);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "RM Restored",
                    Detail = $"Material #{item.IssNo} restored to Material Deductions.",
                    Duration = 3000
                });

                await OnRMRestored.InvokeAsync();
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

        public async Task ContextRestoreRM(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await RestoreRM(ContextRowItem);
            }
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is MakerHiddenRMItem item)
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
