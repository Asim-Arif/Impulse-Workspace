using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Export.Invoices
{
    public partial class PrintInnerLabels : ComponentBase
    {
        [Inject] private IPrintInnerLabelsService PrintInnerLabelsService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = true;

        // Form bindings & lists
        private List<ItemLookupModel> items = new();
        private ItemLookupModel? selectedItem;
        private List<string> orderNumbers = new();
        private string selectedOrderNo = string.Empty;
        private double qty = 0;
        private int numberOfLabels = 1;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                items = await PrintInnerLabelsService.GetActiveItemsAsync();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Load Error", "Failed to load active items: " + ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task<IEnumerable<ItemLookupModel>> SearchItems(string searchText)
        {
            await Task.Yield();
            if (string.IsNullOrWhiteSpace(searchText))
                return items;
            return items.Where(i => i.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                                     i.ItemID.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task OnItemChanged(ItemLookupModel? item)
        {
            selectedItem = item;
            selectedOrderNo = string.Empty;
            orderNumbers.Clear();

            if (selectedItem != null)
            {
                try
                {
                    orderNumbers = await PrintInnerLabelsService.GetOrderNumbersForItemAsync(selectedItem.ItemID);
                    if (orderNumbers.Any())
                    {
                        selectedOrderNo = orderNumbers.First();
                    }
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Data Error", "Failed to load orders for selected item: " + ex.Message);
                }
            }
        }

        private void OnOrderNoChanged(ChangeEventArgs e)
        {
            selectedOrderNo = e.Value?.ToString() ?? string.Empty;
        }

        private async Task OnPrintClick()
        {
            if (selectedItem == null)
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select an Item.");
                return;
            }

            if (numberOfLabels <= 0)
            {
                NotificationServiceManager.ShowWarning("Validation", "Invalid No. of labels.");
                return;
            }

            try
            {
                var request = new ReportRequest
                {
                    ReportName = "ItemLabel.rpt",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@ItemID", selectedItem.ItemID },
                        { "@PONo", selectedOrderNo },
                        { "@No", numberOfLabels },
                        { "@Qty", qty }
                    }
                };

                await ReportNavigation.PrintReportAsync(request);
                NotificationServiceManager.ShowSuccess("Success", "Labels report dispatched successfully.");
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", "Failed to print inner labels: " + ex.Message);
            }
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo("/export");
        }
    }
}
