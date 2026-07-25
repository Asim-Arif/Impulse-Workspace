using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.WebUtilities;
using Impulse.Services;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Pages.Stock.VendGateRcvd
{
    public partial class VendGateRcvd : ComponentBase
    {
        private VendGateRcvdSaveRequest request = new VendGateRcvdSaveRequest();
        
        private object _selectedVendor;
        public object SelectedVendor 
        {
            get => _selectedVendor;
            set
            {
                if (_selectedVendor != value)
                {
                    _selectedVendor = value;
                    OnVendorChanged();
                }
            }
        }
        
        private OrderDropdownItem _selectedOrder;
        public OrderDropdownItem SelectedOrder 
        {
            get => _selectedOrder;
            set
            {
                if (_selectedOrder != value)
                {
                    _selectedOrder = value;
                    _ = OnOrderChangedAsync();
                }
            }
        }

        private List<string> Stores = new List<string>();
        private bool isLoadingItems = false;
        private bool isSaving = false;
        private bool PrintReceivingNote = true;
        
        public bool ShowPicture { get; set; }
        public string SelectedMaterialId { get; set; }

        private void SelectRowMaterial(string materialId)
        {
            SelectedMaterialId = materialId;
        }
        
        private string ReturnUrl;

        protected override async Task OnInitializedAsync()
        {
            var uri = NavigationManager.ToAbsoluteUri(NavigationManager.Uri);
            if (QueryHelpers.ParseQuery(uri.Query).TryGetValue("returnUrl", out var returnUrlStr))
            {
                ReturnUrl = returnUrlStr;
            }
            
            Stores = await _vendGateRcvdService.GetStoresAsync();
        }

        private async Task<IEnumerable<object>> SearchVendors(string searchText)
        {
            var vendors = await _vendGateRcvdService.GetVendorsWithPendingOrdersAsync();
            if (string.IsNullOrEmpty(searchText))
            {
                return vendors;
            }
            return vendors.Where(v => ((dynamic)v).AccTitle.ToLower().Contains(searchText.ToLower())).ToList();
        }

        private async Task<IEnumerable<OrderDropdownItem>> SearchOrders(string searchText)
        {
            if (SelectedVendor == null) return new List<OrderDropdownItem>();
            
            string vendId = ((dynamic)SelectedVendor).AccNo;
            var orders = await _vendGateRcvdService.GetPendingOrdersAsync(vendId);
            
            var items = orders.Select(o => new OrderDropdownItem { OrderText = o }).ToList();
            
            if (string.IsNullOrEmpty(searchText))
            {
                return items;
            }
            return items.Where(o => o.OrderText.ToLower().Contains(searchText.ToLower())).ToList();
        }
        
        private void OnVendorChanged()
        {
            SelectedOrder = null;
            request.ReceivedItems.Clear();
            if (SelectedVendor != null)
            {
                request.VendID = ((dynamic)SelectedVendor).AccNo;
            }
        }

        private async Task OnOrderChangedAsync()
        {
            request.ReceivedItems.Clear();
            if (SelectedOrder != null)
            {
                isLoadingItems = true;
                StateHasChanged();
                
                // Parse order no from "PONo (RunningPONo)" format or just get it from DAC
                // Since our DAC currently uses "SELECT DISTINCT PONo + ' (' + CAST(RunningPONo AS VARCHAR(50)) + ')' AS OrderText, PONo"
                // Oh wait, GetPendingOrdersAsync returns List<string>. 
                // We should make it return an object with both text and PONo.
                // Let's modify OrderDropdownItem logic. The current DAC returns just a list of strings if we mapped it like that, 
                // Wait, in VendGateRcvdDataAccess, GetPendingOrdersAsync returns list of string (OrderText). I will just parse PONo from the string or I can modify the DAC.
                // It's safer to extract the text before the space. e.g. "VPO-24-001 (12)".
                string fullText = SelectedOrder.OrderText;
                string poNo = fullText;
                int spaceIndex = fullText.IndexOf(' ');
                if (spaceIndex > 0)
                {
                    poNo = fullText.Substring(0, spaceIndex);
                }
                
                request.OrderNo = poNo;
                request.ReceivedItems = await _vendGateRcvdService.GetPendingOrderItemsAsync(poNo);
                
                isLoadingItems = false;
                StateHasChanged();
            }
        }

        private async Task Save(bool saveAndNew)
        {
            if (SelectedVendor == null || SelectedOrder == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Warning, Summary = "Validation", Detail = "Please select a Vendor and an Order No." });
                return;
            }
            
            if (!request.ReceivedItems.Any(x => x.QtyRcvd > 0))
            {
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Warning, Summary = "Validation", Detail = "Please enter Qty Rcvd for at least one item." });
                return;
            }
            
            // Validate Store Selection for received items
            if (request.ReceivedItems.Any(x => x.QtyRcvd > 0 && string.IsNullOrWhiteSpace(x.StoreName)))
            {
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Warning, Summary = "Validation", Detail = "Please select a Store Name for all received items." });
                return;
            }

            try
            {
                isSaving = true;
                StateHasChanged();
                
                string rcvId = await _vendGateRcvdService.SaveVendorReceivingAsync(request, "Current_User_Here"); // Replace with actual user context if available
                
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Success, Summary = "Success", Detail = $"Saved Successfully. Rcv ID: {rcvId}" });
                
                if (PrintReceivingNote)
                {
                    var reportReq = new Impulse.Services.ReportRequest
                    {
                        ReportName = "GRN.rpt",
                        SelectionFormula = $"{{VendRcvd.RcvID}}='{rcvId}'"
                    };
                    await ReportNavigationService.PrintReportAsync(reportReq);
                }
                
                if (saveAndNew)
                {
                    request = new VendGateRcvdSaveRequest();
                    SelectedVendor = null;
                    SelectedOrder = null;
                }
                else
                {
                    Close();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Error, Summary = "Error", Detail = ex.Message });
            }
            finally
            {
                isSaving = false;
                StateHasChanged();
            }
        }

        private void Close()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/stock");
            }
        }
        
        public class OrderDropdownItem 
        {
            public string OrderText { get; set; }
        }
    }
}
