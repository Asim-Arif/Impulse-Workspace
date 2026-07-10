using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Radzen;

namespace Impulse.Components.Export.NewCustomInvoiceTabs
{
    public partial class ProformaItemsTab
    {
        [Parameter] public CustomInvoiceViewModel Invoice { get; set; } = default!;
        [Parameter] public string CustCode { get; set; } = string.Empty;
        [Parameter] public string Country { get; set; } = string.Empty;
        [Parameter] public EventCallback<(List<CustomPListModel> Items, string ProformaNo)> OnAddItems { get; set; }

        private List<GenericDropDownModel> CustomerItems = new();
        private List<GenericDropDownModel> ProformaOrders = new();

        private bool ShowDispatchModal = false;
        private long SelectedDispatchListId = 0;
        private List<DispatchListLookupModel> AvailableDispatchLists = new();
        private List<DispatchCartonModel> UnmatchedCartons = new();

        private string _lastCustCode = string.Empty;
        private string _lastCountry = string.Empty;

        private GenericDropDownModel SelectedItem = default!;
        private GenericDropDownModel SelectedProforma = default!;

        private int OrderQty, Balance, Used, Available, InputQtyPerCarton;
        private string InputLotNo = string.Empty;
        private int InputCartonFrom, InputCartonTo;

        private CustomPListModel SelectedProformaData = default!;

        protected override async Task OnParametersSetAsync()
        {
            if (!string.IsNullOrEmpty(CustCode) && !string.IsNullOrEmpty(Country))
            {
                if (CustCode != _lastCustCode || Country != _lastCountry)
                {
                    _lastCustCode = CustCode;
                    _lastCountry = Country;
                    CustomerItems = await InvoiceService.GetCustomerItemsAsync(CustCode, Country);
                    SelectedItem = null;
                    SelectedProforma = null;
                    ProformaOrders.Clear();
                    ResetQuantities();
                }
            }
        }

        private async Task<IEnumerable<GenericDropDownModel>> SearchItems(string searchText)
        {
            if (string.IsNullOrEmpty(searchText))
                return CustomerItems;
            
            return await Task.FromResult(CustomerItems.Where(x => x.DropDownValue_Description.Contains(searchText, StringComparison.InvariantCultureIgnoreCase)));
        }

        private async Task<IEnumerable<GenericDropDownModel>> SearchProformas(string searchText)
        {
            if (string.IsNullOrEmpty(searchText))
                return ProformaOrders;
            
            return await Task.FromResult(ProformaOrders.Where(x => x.DropDownValue_Description.Contains(searchText, StringComparison.InvariantCultureIgnoreCase)));
        }

        private async Task OnItemSelected(GenericDropDownModel item)
        {
            SelectedItem = item;
            ProformaOrders.Clear();
            SelectedProforma = null;
            ResetQuantities();

            if (SelectedItem != null)
            {
                ProformaOrders = await InvoiceService.GetProformasByItemAsync(SelectedItem.DropDownValue_ID, CustCode, Country);
            }
        }

        private async Task OnProformaSelected(GenericDropDownModel proforma)
        {
            SelectedProforma = proforma;
            ResetQuantities();

            if (SelectedProforma != null)
            {
                var proformaData = await InvoiceService.GetItemsForProformaAsync(SelectedProforma.DropDownValue_Description.Split(' ')[0]);
                SelectedProformaData = proformaData.FirstOrDefault(x => x.OrderItemID.ToString() == SelectedProforma.DropDownValue_ID);

                if (SelectedProformaData != null)
                {
                    OrderQty = SelectedProformaData.Qty;
                    Balance = SelectedProformaData.Qty;
                    CalculateUsedAndAvailable();
                }
            }
        }

        private void CalculateUsedAndAvailable()
        {
            if (SelectedProforma == null) return;
            Used = Invoice.PackingListItems.Where(x => x.OrderItemID.ToString() == SelectedProforma.DropDownValue_ID).Sum(x => x.Qty);
            Available = Balance - Used;
        }

        private void ResetQuantities()
        {
            OrderQty = 0;
            Balance = 0;
            Used = 0;
            Available = 0;
            SelectedProformaData = null;
        }

        private async Task AddItem()
        {
            if (InputQtyPerCarton <= 0 || Available <= 0 || SelectedProformaData == null)
                return;

            if (InputQtyPerCarton > Available)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Warning", "Quantity exceeds available amount.");
                return;
            }

            var newItem = new CustomPListModel
            {
                OrderItemID = SelectedProformaData.OrderItemID,
                ProformaInvoice = SelectedProformaData.ProformaInvoice,
                OrderNo = SelectedProformaData.OrderNo,
                ItemCode = SelectedProformaData.ItemCode,
                Description = SelectedProformaData.Description,
                LotNo = InputLotNo,
                CartonFrom = InputCartonFrom,
                CartonTo = InputCartonTo,
                Qty = InputQtyPerCarton
            };

            var itemsToAdd = new List<CustomPListModel> { newItem };
            
            await OnAddItems.InvokeAsync((itemsToAdd, SelectedProformaData.ProformaInvoice));
            
            CalculateUsedAndAvailable();
            InputQtyPerCarton = 0;

            // Notify user
            NotificationService.Notify(NotificationSeverity.Success, "Added", "Item added to invoice.", 2000);
            StateHasChanged();
        }

        private void RemoveItem(CustomPListModel item)
        {
            Invoice.PackingListItems.Remove(item);
            
            // Also need to adjust LineItems. HandleAddItems added to LineItems.
            var lineItem = Invoice.LineItems.FirstOrDefault(x => x.RefID == item.OrderItemID);
            if (lineItem != null)
            {
                lineItem.Qty -= item.Qty;
                if (lineItem.Qty <= 0)
                {
                    Invoice.LineItems.Remove(lineItem);
                }
            }

            CalculateUsedAndAvailable();
            StateHasChanged();
        }

        private async Task HandleQtyKeyPress(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await AddItem();
            }
        }
        private async Task OpenDispatchModal()
        {
            if (string.IsNullOrWhiteSpace(CustCode))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select a customer first.");
                return;
            }
            AvailableDispatchLists = await InvoiceService.GetAvailableDispatchListsAsync(CustCode);
            if (!AvailableDispatchLists.Any())
            {
                NotificationService.Notify(NotificationSeverity.Info, "No Data", "No available dispatch lists found for this customer.");
                return;
            }
            SelectedDispatchListId = 0;
            ShowDispatchModal = true;
        }

        private void CloseDispatchModal()
        {
            ShowDispatchModal = false;
        }

        private async Task ProcessDispatchList()
        {
            if (SelectedDispatchListId == 0) return;

            var cartons = await InvoiceService.GetDispatchListCartonsAsync(SelectedDispatchListId);
            UnmatchedCartons.Clear();
            
            var itemsToAdd = new List<CustomPListModel>();
            string lastProforma = string.Empty;
            Dictionary<long, long> sessionUsedQty = new();

            foreach (var carton in cartons)
            {
                var proformas = await InvoiceService.GetProformaItemsForDispatchMappingAsync(carton.OrderNo, carton.ItemCode);
                long cartonQtyAvailable = carton.CartonQty;

                foreach (var p in proformas)
                {
                    if (cartonQtyAvailable <= 0) break;

                    long dbUsedQty = Invoice.PackingListItems?.Where(i => i.OrderItemID == p.OrderItemID).Sum(i => i.Qty) ?? 0;
                    long thisSessionUsed = sessionUsedQty.ContainsKey(p.OrderItemID) ? sessionUsedQty[p.OrderItemID] : 0;
                    
                    long availableInProforma = p.Qty - (dbUsedQty + thisSessionUsed);
                    
                    if (availableInProforma <= 0) continue;

                    long qtyToUse = Math.Min(cartonQtyAvailable, availableInProforma);
                    cartonQtyAvailable -= qtyToUse;

                    if (sessionUsedQty.ContainsKey(p.OrderItemID))
                        sessionUsedQty[p.OrderItemID] += qtyToUse;
                    else
                        sessionUsedQty[p.OrderItemID] = qtyToUse;

                    var newItem = new CustomPListModel
                    {
                        ID = 0,
                        OrderItemID = p.OrderItemID,
                        ProformaInvoice = p.ProformaInvoice,
                        OrderNo = p.OrderNo,
                        ItemCode = p.ItemCode,
                        Description = p.Description,
                        Qty = (int)qtyToUse,
                        LotNo = carton.LotNo,
                        CartonFrom = (int)carton.CartonNo,
                        CartonTo = (int)carton.CartonNo,
                        CartonQty = (int)carton.CartonNo // From legacy code logic
                    };
                    
                    itemsToAdd.Add(newItem);
                    lastProforma = p.ProformaInvoice;
                }

                if (cartonQtyAvailable > 0)
                {
                    carton.CartonQty = cartonQtyAvailable;
                    UnmatchedCartons.Add(carton);
                }
            }

            if (itemsToAdd.Any())
            {
                await OnAddItems.InvokeAsync((itemsToAdd, lastProforma));
                CalculateUsedAndAvailable();
                NotificationService.Notify(NotificationSeverity.Success, "Success", $"{itemsToAdd.Count} items added from Dispatch List.");
            }
            else if (!UnmatchedCartons.Any())
            {
                NotificationService.Notify(NotificationSeverity.Info, "Notice", "No items could be mapped to active Proformas.");
            }

            ShowDispatchModal = false;
        }
    }
}
