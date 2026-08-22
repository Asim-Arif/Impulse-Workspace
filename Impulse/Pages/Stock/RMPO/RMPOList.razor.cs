using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Impulse.Services;

namespace Impulse.Pages.Stock.RMPO
{
    public partial class RMPOList
    {
        [Parameter]
        [SupplyParameterFromQuery]
        public bool IsSampleList { get; set; } = false;

        private RMPOListSearchFilter Filter { get; set; } = new RMPOListSearchFilter();
        private IEnumerable<RMPOListViewModel> POList { get; set; }
        
        private List<NewVendorViewModel> Vendors = new();
        private List<RawMaterialLookupModel> Materials = new();

        private NewVendorViewModel _selectedVendor;
        private NewVendorViewModel SelectedVendor
        {
            get => _selectedVendor;
            set
            {
                if (_selectedVendor != value)
                {
                    _selectedVendor = value;
                    Filter.VendorID = _selectedVendor?.AccNo;
                }
            }
        }

        private RawMaterialLookupModel _selectedMaterial;
        private RawMaterialLookupModel SelectedMaterial
        {
            get => _selectedMaterial;
            set
            {
                if (_selectedMaterial != value)
                {
                    _selectedMaterial = value;
                    Filter.MaterialID = _selectedMaterial?.RMID1;
                }
            }
        }

        private bool IsLoading { get; set; } = false;
        
        [Inject]
        private IJSRuntime JS { get; set; }

        [Inject]
        private BlazorContextMenu.IBlazorContextMenuService BlazorContextMenuService { get; set; }

        protected override async Task OnInitializedAsync()
        {
            Filter.DateRangeIndex = 0; // Last 15 Days
            Filter.ShowSamplePOs = IsSampleList;
            UpdateDateRange();

            Vendors = (await VendorService.GetVendorsAsync()).ToList();
            Materials = (await RMService.GetAvailableRawMaterialsAsync()).ToList();

            await LoadData();
        }

        private void OnDateRangeChanged()
        {
            UpdateDateRange();
        }

        private void UpdateDateRange()
        {
            Filter.DateTo = DateTime.Today;
            switch (Filter.DateRangeIndex)
            {
                case 0: Filter.DateFrom = DateTime.Today.AddDays(-15); break;
                case 1: Filter.DateFrom = DateTime.Today.AddDays(-30); break;
                case 2: Filter.DateFrom = DateTime.Today.AddDays(-60); break;
                case 3: Filter.DateFrom = DateTime.Today.AddDays(-90); break;
                case 4: Filter.DateFrom = DateTime.Today.AddYears(-1); break;
                case 5: // Custom Date - Leave as is
                    break;
            }
        }

        private async Task<IEnumerable<NewVendorViewModel>> SearchVendors(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Vendors;
            return await Task.FromResult(Vendors.Where(x => x.VendorName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<RawMaterialLookupModel>> SearchMaterials(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Materials;
            return await Task.FromResult(Materials.Where(x => x.RMName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnRefreshClicked()
        {
            await LoadData();
        }

        private async Task LoadData()
        {
            IsLoading = true;
            try
            {
                POList = await VendOrderService.GetRMPOListAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Data", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void EditPO(RMPOListViewModel item)
        {
            NavigationManager.NavigateTo($"/stock/new-rm-po?OrderNo={item.OrderNo}&returnUrl=/stock/rmpolist");
        }

        private async Task OnEditPO(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
            {
                if (item.Final == true)
                {
                    NotificationService.ShowWarning("Warning", "PO is closed, Can't Edit.");
                    return;
                }
                EditPO(item);
            }
        }

        private async Task OnDeletePO(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
            {
                if (item.Final == true)
                {
                    NotificationService.ShowWarning("Warning", "PO is closed, Can't Delete.");
                    return;
                }

                bool confirmed = await JS.InvokeAsync<bool>("confirm", "Are you sure you want to delete this order?");
                if (confirmed)
                {
                    try
                    {
                        await VendOrderService.DeleteVendOrderAsync(item.OrderNo);
                        NotificationService.ShowSuccess("Success", "Order deleted.");
                        await LoadData();
                    }
                    catch (Exception ex)
                    {
                        NotificationService.ShowError("Error", ex.Message);
                    }
                }
            }
        }

        private async Task OnEndPO(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
            {
                bool confirmed = await JS.InvokeAsync<bool>("confirm", "Are you sure to end this purchase order?");
                if (confirmed)
                {
                    try
                    {
                        var order = await VendOrderService.GetVendOrderAsync(item.OrderNo);
                        if (order != null)
                        {
                            order.Final = true;
                            await VendOrderService.SaveVendOrderAsync(order);
                            NotificationService.ShowSuccess("Success", "PO closed successfully.");
                            await LoadData();
                        }
                    }
                    catch (Exception ex)
                    {
                        NotificationService.ShowError("Error", ex.Message);
                    }
                }
            }
        }

        private async Task ShowOptionsMenu(Microsoft.AspNetCore.Components.Web.MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        // --- Reporting ---
        private async Task OnPrintOrder(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
                await ReportNavigation.PrintReportAsync(new ReportRequest { ReportName = "RMPO.rpt", SelectionFormula = $"{{VendOrders.OrderNo}}='{item.OrderNo}'" });
        }

        private async Task OnPrintOrderStatus(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
                await ReportNavigation.PrintReportAsync(new ReportRequest { ReportName = "RMPOWithReceiving.rpt", SelectionFormula = $"{{VendOrders.OrderNo}}='{item.OrderNo}'" });
        }

        private async Task OnPrintOrderZeroRate(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
                await ReportNavigation.PrintReportAsync(new ReportRequest 
                { 
                    ReportName = "RMPOWP.rpt", 
                    SelectionFormula = $"{{VendOrders.OrderNo}}='{item.OrderNo}'",
                    FormulaValues = new Dictionary<string, object> { { "ZeroRate", true } }
                });
        }

        private async Task OnPrintOrderOfficeCopy(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
                await ReportNavigation.PrintReportAsync(new ReportRequest { ReportName = "RMPOOfficeCopy.rpt", SelectionFormula = $"{{VendOrders.OrderNo}}='{item.OrderNo}'" });
        }

        private async Task OnPrintOrderVendorCopy(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
                await ReportNavigation.PrintReportAsync(new ReportRequest { ReportName = "RMPOVendorCopy.rpt", SelectionFormula = $"{{VendOrders.OrderNo}}='{item.OrderNo}'" });
        }

        private async Task OnPrintOrderAccountsCopy(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
                await ReportNavigation.PrintReportAsync(new ReportRequest { ReportName = "RMPOAccountCopy.rpt", SelectionFormula = $"{{VendOrders.OrderNo}}='{item.OrderNo}'" });
        }

        private async Task OnPrintOrderListSimple(BlazorContextMenu.ItemClickEventArgs e)
        {
            string selectionFormula = $"{{VVendOrders.DT}} in Date({Filter.DateFrom?.Year ?? 1900}, {Filter.DateFrom?.Month ?? 1}, {Filter.DateFrom?.Day ?? 1}) to Date({Filter.DateTo?.Year ?? 2099}, {Filter.DateTo?.Month ?? 1}, {Filter.DateTo?.Day ?? 1})";
            
            await ReportNavigation.PrintReportAsync(new ReportRequest { ReportName = "rptVendPOListDateWise.rpt", SelectionFormula = selectionFormula });
        }
        
        // --- PDF Attachment ---
        private RMPOListViewModel selectedItemForPdf;
        
        private async Task TriggerFileUpload(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is RMPOListViewModel item)
            {
                selectedItemForPdf = item;
                await JS.InvokeVoidAsync("eval", "document.getElementById('pdfFileInput').click()");
            }
        }

        private async Task OpenPDF(RMPOListViewModel item)
        {
            if (item.PDFAttachment != null && item.PDFAttachment.Length > 0)
            {
                var base64 = Convert.ToBase64String(item.PDFAttachment);
                await JS.InvokeVoidAsync("window.openPdfFromBase64", base64);
            }
        }

        private async Task OnPdfFileSelected(InputFileChangeEventArgs e)
        {
            if (selectedItemForPdf == null) return;
            
            var file = e.File;
            if (file != null && file.ContentType == "application/pdf")
            {
                try
                {
                    using var stream = file.OpenReadStream(maxAllowedSize: 10485760); // 10 MB max
                    using var ms = new System.IO.MemoryStream();
                    await stream.CopyToAsync(ms);
                    
                    await VendOrderService.UpdatePDFAttachmentAsync(selectedItemForPdf.OrderNo, ms.ToArray());
                    NotificationService.ShowSuccess("Success", "PDF attached successfully.");
                    await LoadData();
                }
                catch (Exception ex)
                {
                    NotificationService.ShowError("Error", "Failed to upload PDF: " + ex.Message);
                }
            }
            else
            {
                NotificationService.ShowError("Error", "Please select a valid PDF file.");
            }
        }
    }
}
