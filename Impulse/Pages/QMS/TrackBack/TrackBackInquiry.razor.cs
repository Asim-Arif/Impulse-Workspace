using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Components.Common;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.TrackBack
{
    public partial class TrackBackInquiry : ComponentBase
    {
        [Inject] private ITrackBackDataAccess TrackBackDataAccess { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;

        private string? selectedInvoice;
        private List<string> recentInvoices = new();
        private List<TrackBackInvoiceItemDto> invoiceItems = new();
        private TrackBackInvoiceItemDto? selectedItem;

        private List<TrackBackLotDto> lots = new();
        private List<TrackBackOrderDto> orders = new();
        private List<TrackBackMasterPODto> masterPOs = new();

        private string itemSearchFilter = string.Empty;
        private bool isLoadingItems = false;
        private bool isLoadingDetails = false;

        private AttachmentModal? attachmentModalRef;

        private List<TrackBackInvoiceItemDto> filteredItems
        {
            get
            {
                if (string.IsNullOrWhiteSpace(itemSearchFilter))
                {
                    return invoiceItems;
                }

                var q = itemSearchFilter.Trim();
                return invoiceItems.Where(i =>
                    (!string.IsNullOrEmpty(i.CompItemCode) && i.CompItemCode.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(i.ItemName) && i.ItemName.Contains(q, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadRecentInvoices();
        }

        private async Task LoadRecentInvoices()
        {
            try
            {
                recentInvoices = await TrackBackDataAccess.GetRecentCustomInvoicesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load invoices: {ex.Message}");
            }
        }

        private async Task<IEnumerable<string>> SearchInvoices(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return recentInvoices.Take(100);
            }

            var results = await TrackBackDataAccess.GetRecentCustomInvoicesAsync(searchText);
            return results;
        }

        private async Task OnInvoiceChanged(string? invoiceNo)
        {
            selectedInvoice = invoiceNo;
            if (!string.IsNullOrWhiteSpace(selectedInvoice))
            {
                await LoadInvoiceItems();
            }
            else
            {
                ClearDetails();
            }
        }

        private async Task LoadInvoiceItems()
        {
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;

            try
            {
                isLoadingItems = true;
                ClearDetails();

                invoiceItems = await TrackBackDataAccess.GetInvoiceItemsAsync(selectedInvoice);

                if (invoiceItems.Any())
                {
                    await SelectItem(invoiceItems.First());
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Info, "No Items", $"No items found for invoice '{selectedInvoice}'.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load invoice items: {ex.Message}");
            }
            finally
            {
                isLoadingItems = false;
            }
        }

        private async Task SelectItem(TrackBackInvoiceItemDto item)
        {
            selectedItem = item;
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;

            try
            {
                isLoadingDetails = true;

                var lotsTask = TrackBackDataAccess.GetLotsByItemAsync(selectedInvoice, item.CompItemCode);
                var ordersTask = TrackBackDataAccess.GetOrdersByItemAsync(selectedInvoice, item.CompItemCode);
                var posTask = TrackBackDataAccess.GetMasterPOsByItemAsync(selectedInvoice, item.CompItemCode);

                await Task.WhenAll(lotsTask, ordersTask, posTask);

                lots = lotsTask.Result;
                orders = ordersTask.Result;
                masterPOs = posTask.Result;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load item breakdown: {ex.Message}");
            }
            finally
            {
                isLoadingDetails = false;
            }
        }

        private void ClearDetails()
        {
            selectedItem = null;
            lots.Clear();
            orders.Clear();
            masterPOs.Clear();
        }

        private void OnItemFilterInput(ChangeEventArgs e)
        {
            itemSearchFilter = e.Value?.ToString() ?? string.Empty;
        }

        private void NavigateToHub()
        {
            NavigationManager.NavigateTo("/qms");
        }

        private void OpenLotAttachment(string lotNo)
        {
            attachmentModalRef?.Open(
                entityType: "Lot",
                entityId: lotNo,
                secondaryId: selectedInvoice,
                displayName: $"Lot: {lotNo}"
            );
        }

        private async Task ShowOptionsMenu(MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        #region Report Handlers

        private async Task PrintDHR()
        {
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptDHR.rpt",
                    SelectionFormula = $"{{VRptDHR.CustomInvoice}} = '{selectedInvoice}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print DHR: {ex.Message}");
            }
        }

        private async Task PrintInvoiceItems()
        {
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Invoice_Items.rpt",
                    SelectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}} = '{selectedInvoice}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print invoice items: {ex.Message}");
            }
        }

        private async Task PrintItemLabels()
        {
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "TBI_Item_Labels.rpt",
                    SelectionFormula = $"{{VCustomInvoiceItems.CustomInvoice}} = '{selectedInvoice}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print item labels: {ex.Message}");
            }
        }

        private async Task PrintTechnicalDrawings()
        {
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "TBI_Item_TechDraws.rpt",
                    SelectionFormula = $"{{VCustomInvoiceItems.CustomInvoice}} = '{selectedInvoice}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print technical drawings: {ex.Message}");
            }
        }

        private async Task PrintProductionOrders()
        {
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Invoice_Order_Items.rpt",
                    SelectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}} = '{selectedInvoice}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print production orders: {ex.Message}");
            }
        }

        private async Task PrintPurchaseOrders()
        {
            if (string.IsNullOrWhiteSpace(selectedInvoice)) return;

            try
            {
                var poList = await TrackBackDataAccess.GetInvoiceMasterPONosAsync(selectedInvoice);
                if (!poList.Any())
                {
                    NotificationService.Notify(NotificationSeverity.Info, "Information", "No Master Purchase Orders found for this invoice.");
                    return;
                }

                string formattedPOs = string.Join(",", poList.Select(p => $"'{p}'"));
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "IssList.rpt",
                    SelectionFormula = $"{{VendIssued.MasterPONo}} IN [{formattedPOs}]"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print purchase orders: {ex.Message}");
            }
        }

        private async Task PrintPTC(string lotNo)
        {
            if (string.IsNullOrWhiteSpace(lotNo)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "PTCQEL.rpt",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@LotNo", lotNo }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print PTC: {ex.Message}");
            }
        }

        private async Task PrintMasterPOStatus(string masterPONo)
        {
            if (string.IsNullOrWhiteSpace(masterPONo)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "MasterPOStatus.rpt",
                    SelectionFormula = $"{{MasterPOs.MasterPONo}} = '{masterPONo}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print Master PO Status: {ex.Message}");
            }
        }

        private async Task PrintSingleOrder(string orderNo)
        {
            if (string.IsNullOrWhiteSpace(orderNo)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "rptProductionOrderWrangler.rpt",
                    SelectionFormula = $"{{CustOrders.OrderNo}} = '{orderNo}'",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@OrderNo", orderNo }
                    }
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print production order: {ex.Message}");
            }
        }

        #endregion
    }
}

