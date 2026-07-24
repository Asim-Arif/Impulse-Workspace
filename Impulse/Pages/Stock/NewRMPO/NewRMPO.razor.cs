using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.JSInterop;
using Radzen;
using System;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.NewRMPO
{
    public partial class NewRMPO : ComponentBase
    {
        [Inject]
        private IVendOrderService VendOrderService { get; set; }

        [Inject]
        private NavigationManager NavigationManager { get; set; }

        [Inject]
        private NotificationService NotificationService { get; set; }

        public VendOrderViewModel Order { get; set; } = new VendOrderViewModel();
        public bool IsSaving { get; set; }

        [Parameter]
        [SupplyParameterFromQuery]
        public string OrderNo { get; set; }

        protected override async Task OnInitializedAsync()
        {
            if (!string.IsNullOrEmpty(OrderNo))
            {
                Order = await VendOrderService.GetVendOrderAsync(OrderNo);
                if (Order == null || string.IsNullOrEmpty(Order.OrderNo))
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Error", "Order not found");
                    NavigationManager.NavigateTo("/stock/new-rm-po");
                    Order = new VendOrderViewModel();
                }
            }
            else
            {
                Order = new VendOrderViewModel
                {
                    DT = DateTime.Now,
                    DeliveryDT = DateTime.Now,
                    POType = 0 // Default to Raw Material
                };
            }
        }

        private async Task SaveOrder()
        {
            try
            {
                IsSaving = true;

                if (string.IsNullOrEmpty(Order.VendID))
                {
                    NotificationService.Notify(NotificationSeverity.Warning, "Validation Error", "Please select a Vendor.");
                    return;
                }

                if (Order.OrderDetails == null || Order.OrderDetails.Count == 0)
                {
                    NotificationService.Notify(NotificationSeverity.Warning, "Validation Error", "Order must contain at least one material.");
                    return;
                }

                string result = await VendOrderService.SaveVendOrderAsync(Order);
                NotificationService.Notify(NotificationSeverity.Success, "Success", "Purchase Order saved successfully.");
                
                NavigationManager.NavigateTo("/stock/new-rm-po?OrderNo=" + result, forceLoad: true);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to save: {ex.Message}");
            }
            finally
            {
                IsSaving = false;
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo("/"); // Assuming dashboard or list view
        }
    }
}
