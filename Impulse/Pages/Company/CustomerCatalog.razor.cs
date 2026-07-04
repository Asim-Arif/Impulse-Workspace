using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Company;

namespace Impulse.Pages.Company
{
    public partial class CustomerCatalog : ComponentBase
    {
        [Inject] private ICustomerCatalogService CustomerCatalogService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private IJSRuntime JSRuntime { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = false;
        private bool isGridLoading = false;
        private string searchText = string.Empty;

        // Lookups — page filters
        private List<CustomerLookupModel> customers = new();
        private List<string> countries = new();
        private List<ItemGroupLookupModel> itemGroups = new();
        private List<ItemLookupModel> unassignedItems = new();

        // Lookups — edit modal
        private List<string> units = new();
        private List<FinQualityLookupModel> finQualityList = new();
        private List<string> mdGroups = new();

        // Selection Filters
        private CustomerLookupModel? selectedCustomer = null;
        private string selectedCountry = string.Empty;
        private ItemGroupLookupModel? selectedItemGroup = null;
        private ItemLookupModel? selectedUnassignedItem = null;

        // Details
        private double maxDiscount = 0.0;
        private bool showTopQuality = false;
        private bool showPicture = false;

        // Grid items
        private List<CustomerCatalogItemViewModel> catalogItems = new();
        private CustomerCatalogItemViewModel? selectedCatalogItem = null;
        private string selectedItemPicBase64 = string.Empty;

        // Modal States: Rate Adjustment
        private bool showRateAdjustmentModal = false;
        private double rateAdjustmentPercentage = 0.0;
        private bool rateAdjustmentIsIncrease = true;

        // Modal States: Edit catalog item
        private bool showEditModal = false;
        private int activeEditTab = 0; // 0 = Device Info, 1 = Other Info, 2 = Barcode
        private CustomerCatalogItemViewModel editingItem = new();
        private List<CustomerCatalogRevisionViewModel> revisions = new();
        private CustomerCatalogRevisionViewModel? selectedRevision = null; // <Current> or a historical one
        private bool isShowingHistoricPrices = false;

        // Tab 3 file handling
        private string barcodeFilePath = string.Empty;         // displayable path name
        private byte[]? artworkPDFBytes = null;               // actual bytes for upload
        private string artworkPDFFileName = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                customers = await CustomerCatalogService.GetCustomersAsync();

                var rawGroups = await CustomerCatalogService.GetItemGroupsAsync();
                itemGroups = new List<ItemGroupLookupModel> { new ItemGroupLookupModel { ID = 0, Description = "<All>" } };
                itemGroups.AddRange(rawGroups);

                // Load edit-modal lookups once
                units = await CustomerCatalogService.GetUnitsAsync();
                finQualityList = await CustomerCatalogService.GetFinishingQualitiesAsync();
                mdGroups = await CustomerCatalogService.GetMDGroupsAsync();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Initialization Error", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task OnCustomerSelected(CustomerLookupModel? cust)
        {
            selectedCustomer = cust;
            countries.Clear();
            selectedCountry = string.Empty;
            maxDiscount = 0.0;
            catalogItems.Clear();
            selectedCatalogItem = null;
            selectedItemPicBase64 = string.Empty;
            unassignedItems.Clear();
            selectedUnassignedItem = null;

            if (cust != null)
            {
                countries = await CustomerCatalogService.GetCountriesForCustomerAsync(cust.CustCode);
                if (countries.Any())
                {
                    selectedCountry = countries.First();
                    await LoadCustomerCountryData();
                }
            }
        }

        private async Task OnCountrySelected(string? country)
        {
            selectedCountry = country ?? string.Empty;
            await LoadCustomerCountryData();
        }

        private async Task OnItemGroupSelected(ItemGroupLookupModel? group)
        {
            selectedItemGroup = group;
            await RefreshGrid();
        }

        private async Task LoadCustomerCountryData()
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry)) return;

            isGridLoading = true;
            try
            {
                maxDiscount = await CustomerCatalogService.GetMaxDiscountAsync(selectedCustomer.CustCode, selectedCountry);
                unassignedItems = await CustomerCatalogService.GetUnassignedItemsAsync(selectedCustomer.CustCode, selectedCountry);
                selectedUnassignedItem = null;
                await RefreshGrid();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Loading Data", ex.Message);
            }
            finally
            {
                isGridLoading = false;
            }
        }

        private async Task RefreshGrid()
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry)) return;

            isGridLoading = true;
            try
            {
                int? groupId = selectedItemGroup?.ID == 0 ? null : selectedItemGroup?.ID;
                catalogItems = await CustomerCatalogService.GetCatalogItemsAsync(
                    selectedCustomer.CustCode,
                    selectedCountry,
                    groupId);

                if (selectedCatalogItem != null)
                {
                    var updated = catalogItems.FirstOrDefault(x => x.EntryID == selectedCatalogItem.EntryID);
                    selectedCatalogItem = updated;
                    if (updated != null && showPicture)
                        await LoadItemPicture(updated.CompItemID);
                    else
                        selectedItemPicBase64 = string.Empty;
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Grid Refresh Error", ex.Message);
            }
            finally
            {
                isGridLoading = false;
            }
        }

        private IEnumerable<CustomerCatalogItemViewModel> GetFilteredItems()
        {
            if (string.IsNullOrWhiteSpace(searchText)) return catalogItems;
            return catalogItems.Where(x =>
                x.ItemID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                x.CompItemID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                (x.Description != null && x.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (x.OCR != null && x.OCR.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (x.OCR2 != null && x.OCR2.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task HighlightItem(CustomerCatalogItemViewModel item)
        {
            selectedCatalogItem = item;
            if (showPicture)
                await LoadItemPicture(item.CompItemID);
            else
                selectedItemPicBase64 = string.Empty;
        }

        private async Task LoadItemPicture(string compItemId)
        {
            try
            {
                var bytes = await CustomerCatalogService.GetItemPicAsync(compItemId);
                selectedItemPicBase64 = bytes?.Length > 0 ? Convert.ToBase64String(bytes) : string.Empty;
            }
            catch
            {
                selectedItemPicBase64 = string.Empty;
            }
        }

        private async Task OnShowPictureChanged(ChangeEventArgs e)
        {
            showPicture = (bool)(e.Value ?? false);
            if (showPicture && selectedCatalogItem != null)
                await LoadItemPicture(selectedCatalogItem.CompItemID);
            else
                selectedItemPicBase64 = string.Empty;
        }

        private async Task OnShowTopQualityChanged(ChangeEventArgs e)
        {
            showTopQuality = (bool)(e.Value ?? false);
            await RefreshGrid();
        }

        // ── Item Assignment ──────────────────────────────────────────────────

        private async Task AssignSelectedItem()
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry)) return;
            if (selectedUnassignedItem == null)
            {
                NotificationServiceManager.ShowWarning("Warning", "Please select an item to assign.");
                return;
            }
            try
            {
                await CustomerCatalogService.AssignItemAsync(selectedCustomer.CustCode, selectedCountry, selectedUnassignedItem.ItemID);
                NotificationServiceManager.ShowSuccess("Assigned", $"Item '{selectedUnassignedItem.ItemID}' assigned successfully.");
                await LoadCustomerCountryData();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Assignment Error", ex.Message);
            }
        }

        private async Task AssignAllItems()
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry)) return;
            bool confirm = await NotificationServiceManager.ShowQuestionNotification("Assign All", "Are you sure you want to assign all unassigned items?");
            if (!confirm) return;
            try
            {
                await CustomerCatalogService.AssignAllItemsAsync(selectedCustomer.CustCode, selectedCountry);
                NotificationServiceManager.ShowSuccess("Assigned All", "All unassigned items successfully mapped.");
                await LoadCustomerCountryData();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Assign All Error", ex.Message);
            }
        }

        private async Task UnassignCatalogItem(CustomerCatalogItemViewModel item)
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry)) return;
            try
            {
                string? orderNo = await CustomerCatalogService.ItemIsInOrderAsync(item.CompItemID, selectedCustomer.CustCode, selectedCountry);
                if (!string.IsNullOrEmpty(orderNo))
                {
                    NotificationServiceManager.ShowWarning("In Use", $"Item is used in Order '{orderNo}'. Cannot delete.");
                    return;
                }
                bool confirm = await NotificationServiceManager.ShowQuestionNotification("Remove Item", $"Remove '{item.ItemID}' from catalog?");
                if (!confirm) return;

                await CustomerCatalogService.UnassignItemAsync(selectedCustomer.CustCode, selectedCountry, item.CompItemID);
                NotificationServiceManager.ShowSuccess("Removed", "Item removed from catalog.");
                await LoadCustomerCountryData();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Removal Error", ex.Message);
            }
        }

        private async Task UnassignAllItems()
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry)) return;
            bool confirm = await NotificationServiceManager.ShowQuestionNotification("Remove All", "Unassign all unused catalog items? Items in active orders will not be removed.");
            if (!confirm) return;
            try
            {
                await CustomerCatalogService.UnassignAllItemsAsync(selectedCustomer.CustCode, selectedCountry);
                NotificationServiceManager.ShowSuccess("Cleanup Complete", "All unused items unassigned.");
                await LoadCustomerCountryData();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Cleanup Error", ex.Message);
            }
        }

        // ── Context Menu ─────────────────────────────────────────────────────

        private async Task OpenRowMenu(MouseEventArgs e, CustomerCatalogItemViewModel item)
        {
            selectedCatalogItem = item;
            if (showPicture)
                await LoadItemPicture(item.CompItemID);
            await BlazorContextMenuService.ShowMenu("catalogRowMenu", (int)e.ClientX, (int)e.ClientY + 15, item);
        }

        // ── Rate Adjustment Modal ────────────────────────────────────────────

        private void TriggerRateAdjustment(CustomerCatalogItemViewModel item, bool isIncrease)
        {
            rateAdjustmentIsIncrease = isIncrease;
            rateAdjustmentPercentage = 0.0;
            showRateAdjustmentModal = true;
        }

        private async Task SaveRateAdjustment()
        {
            if (selectedCustomer == null) return;
            if (rateAdjustmentPercentage <= 0.0)
            {
                NotificationServiceManager.ShowWarning("Warning", "Please enter a percentage greater than 0.");
                return;
            }
            try
            {
                await CustomerCatalogService.UpdateRatesAsync(selectedCustomer.CustCode, rateAdjustmentPercentage, rateAdjustmentIsIncrease);
                NotificationServiceManager.ShowSuccess("Rates Updated", $"Rates {(rateAdjustmentIsIncrease ? "increased" : "decreased")} by {rateAdjustmentPercentage}% for all countries of this customer.");
                showRateAdjustmentModal = false;
                await RefreshGrid();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Rate Update Error", ex.Message);
            }
        }

        // ── Edit Modal ───────────────────────────────────────────────────────

        private async Task OpenEditModal(CustomerCatalogItemViewModel item)
        {
            selectedCatalogItem = item;
            activeEditTab = 0;
            isShowingHistoricPrices = false;
            artworkPDFBytes = null;
            artworkPDFFileName = string.Empty;

            editingItem = new CustomerCatalogItemViewModel
            {
                EntryID = item.EntryID,
                CustCode = item.CustCode,
                Country = item.Country,
                ItemID = item.ItemID,
                Unit = item.Unit,
                CompItemID = item.CompItemID,
                Description = item.Description,
                FOB = item.FOB,
                ExWorks = item.ExWorks,
                CnFAir = item.CnFAir,
                CnFSea = item.CnFSea,
                CIFAir = item.CIFAir,
                CIFSea = item.CIFSea,
                CnIAir = item.CnIAir,
                CnISea = item.CnISea,
                PackingMode = item.PackingMode,
                BarcodeNo = item.BarcodeNo,
                SpecialInstructions = item.SpecialInstructions,
                StampInstructions = item.StampInstructions,
                PackingInstructions = item.PackingInstructions,
                ItemColor = item.ItemColor,
                BarcodeFile = item.BarcodeFile,
                FOBTop = item.FOBTop,
                ExWorksTop = item.ExWorksTop,
                CnFAirTop = item.CnFAirTop,
                CnFSeaTop = item.CnFSeaTop,
                CIFAirTop = item.CIFAirTop,
                CIFSeaTop = item.CIFSeaTop,
                CnIAirTop = item.CnIAirTop,
                CnISeaTop = item.CnISeaTop,
                ILO = item.ILO,
                OCR = item.OCR,
                OCR2 = item.OCR2,
                FinQuality = item.FinQuality,
                Temper_Rate = item.Temper_Rate,
                First_Inspection_Rate = item.First_Inspection_Rate,
                MDMA = item.MDMA,
                SFDA_Listing_No = item.SFDA_Listing_No,
                MD_Group = item.MD_Group,
                PageNo = item.PageNo,
                CompDesc = item.CompDesc
            };

            barcodeFilePath = item.BarcodeFile ?? string.Empty;

            // PageNo is not in the view — fetch directly from base table
            try
            {
                editingItem.PageNo = await CustomerCatalogService.GetPageNoAsync(item.EntryID);
            }
            catch
            {
                editingItem.PageNo = string.Empty;
            }

            try
            {
                revisions = await CustomerCatalogService.GetRevisionsAsync(item.EntryID);
            }
            catch
            {
                revisions = new List<CustomerCatalogRevisionViewModel>();
            }

            selectedRevision = null; // <Current Prices>
            showEditModal = true;
        }

        private void SelectRevision(CustomerCatalogRevisionViewModel? rev)
        {
            selectedRevision = rev;
            if (rev == null)
            {
                // Restore current prices from the catalog item
                editingItem.FOB = selectedCatalogItem!.FOB;
                editingItem.ExWorks = selectedCatalogItem!.ExWorks;
                editingItem.CnISea = selectedCatalogItem!.CnISea;
                editingItem.CnIAir = selectedCatalogItem!.CnIAir;
                editingItem.CIFSea = selectedCatalogItem!.CIFSea;
                editingItem.CIFAir = selectedCatalogItem!.CIFAir;
                editingItem.CnFSea = selectedCatalogItem!.CnFSea;
                editingItem.CnFAir = selectedCatalogItem!.CnFAir;
                editingItem.FOBTop = selectedCatalogItem!.FOBTop;
                editingItem.ExWorksTop = selectedCatalogItem!.ExWorksTop;
                editingItem.CnISeaTop = selectedCatalogItem!.CnISeaTop;
                editingItem.CnIAirTop = selectedCatalogItem!.CnIAirTop;
                editingItem.CIFSeaTop = selectedCatalogItem!.CIFSeaTop;
                editingItem.CIFAirTop = selectedCatalogItem!.CIFAirTop;
                editingItem.CnFSeaTop = selectedCatalogItem!.CnFSeaTop;
                editingItem.CnFAirTop = selectedCatalogItem!.CnFAirTop;
                editingItem.Description = selectedCatalogItem!.Description;
                isShowingHistoricPrices = false;
            }
            else
            {
                // Load historic prices — read-only view (matches VB6 lstRev_Click)
                editingItem.FOB = rev.FOB;
                editingItem.ExWorks = rev.FOBSialkot;    // FOBSialkot = ExWorks
                editingItem.CnISea = rev.CnISea;
                editingItem.CnIAir = rev.CnIAir;
                editingItem.CIFSea = rev.CIFSea;
                editingItem.CIFAir = rev.CIFAir;
                editingItem.CnFSea = rev.CnFSea;
                editingItem.CnFAir = rev.CnFAir;
                editingItem.FOBTop = rev.FOBTop;
                editingItem.ExWorksTop = rev.ExWorksTop;
                editingItem.CnISeaTop = rev.CnISeaTop;
                editingItem.CnIAirTop = rev.CnIAirTop;
                editingItem.CIFSeaTop = rev.CIFSeaTop;
                editingItem.CIFAirTop = rev.CIFAirTop;
                editingItem.CnFSeaTop = rev.CnFSeaTop;
                editingItem.CnFAirTop = rev.CnFAirTop;
                editingItem.Description = rev.Description;
                isShowingHistoricPrices = true;
            }
        }

        private async Task SaveItemEdits()
        {
            // Validate required fields (matching VB6 SaveItem)
            if (string.IsNullOrWhiteSpace(editingItem.Unit))
            {
                NotificationServiceManager.ShowWarning("Validation", "No Unit selected.");
                return;
            }
            if (string.IsNullOrWhiteSpace(editingItem.ItemID))
            {
                NotificationServiceManager.ShowWarning("Validation", "Invalid Customer Ref. Code.");
                return;
            }

            // Copy barcodeFilePath back into model
            editingItem.BarcodeFile = barcodeFilePath;

            try
            {
                bool success = await CustomerCatalogService.UpdateCatalogItemDetailAsync(editingItem);
                if (!success)
                {
                    NotificationServiceManager.ShowError("Save Error", "Failed to update catalog details.");
                    return;
                }

                // Upload PDF if one was selected
                if (artworkPDFBytes != null && artworkPDFBytes.Length > 0)
                {
                    await CustomerCatalogService.UploadArtworkPDFAsync(editingItem.EntryID, artworkPDFBytes);
                }

                NotificationServiceManager.ShowSuccess("Success", "Catalog details updated successfully.");
                showEditModal = false;
                await RefreshGrid();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Save Error", ex.Message);
            }
        }

        // ── File Browse Handlers ─────────────────────────────────────────────

        private async Task OnBarcodeFileSelected(InputFileChangeEventArgs e)
        {
            var file = e.File;
            barcodeFilePath = file.Name;
        }

        private async Task OnArtworkPDFSelected(InputFileChangeEventArgs e)
        {
            var file = e.File;
            artworkPDFFileName = file.Name;
            using var ms = new System.IO.MemoryStream();
            await file.OpenReadStream(maxAllowedSize: 10 * 1024 * 1024).CopyToAsync(ms);
            artworkPDFBytes = ms.ToArray();
        }

        // ── Typeahead Search Methods ─────────────────────────────────────────

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return customers;
            return await Task.FromResult(customers.Where(c =>
                c.Name.Contains(text, StringComparison.OrdinalIgnoreCase) ||
                c.CustCode.Contains(text, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<string>> SearchCountries(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return countries;
            return await Task.FromResult(countries.Where(c => c.Contains(text, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<ItemGroupLookupModel>> SearchItemGroups(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return itemGroups;
            return await Task.FromResult(itemGroups.Where(g =>
                g.Description.Contains(text, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<CustomerCatalogItemViewModel>> SearchCatalogItems(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return catalogItems;
            return await Task.FromResult(catalogItems.Where(x =>
                x.ItemID.Contains(text, StringComparison.OrdinalIgnoreCase) ||
                (x.Description != null && x.Description.Contains(text, StringComparison.OrdinalIgnoreCase))
            ).ToList());
        }

        private async Task<IEnumerable<ItemLookupModel>> SearchUnassignedItems(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return unassignedItems;
            return await Task.FromResult(unassignedItems.Where(x =>
                x.ItemID.Contains(text, StringComparison.OrdinalIgnoreCase) ||
                x.ItemName.Contains(text, StringComparison.OrdinalIgnoreCase)
            ).ToList());
        }

        private async Task<IEnumerable<string>> SearchUnits(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return units;
            return await Task.FromResult(units.Where(u => u.Contains(text, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnItemSearchSelected(CustomerCatalogItemViewModel? item)
        {
            if (item != null) await HighlightItem(item);
        }

        // ── Reporting ────────────────────────────────────────────────────────

        private async Task PrintPriceList(bool bTop = false)
        {
            if (selectedCustomer == null || string.IsNullOrEmpty(selectedCountry)) return;
            string? dateStr = await JSRuntime.InvokeAsync<string>("prompt", "Please Enter Validity Date (dd/MM/yyyy)", DateTime.Today.AddMonths(1).ToString("dd/MM/yyyy"));
            if (string.IsNullOrEmpty(dateStr)) return;
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Company.CustPriceList,
                    SelectionFormula = $"{{ForeignCustomers.CustCode}}='{selectedCustomer.CustCode}'",
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "TopPrices", bTop ? "True" : "False" },
                        { "ValidityDT", $"#{dateStr}#" }
                    }
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintCatalog(bool detailed)
        {
            if (selectedCustomer == null) return;
            try
            {
                var request = new ReportRequest
                {
                    ReportName = detailed ? ReportNames.Company.CustomerCatalogWithDetails : ReportNames.Company.CustomerCatalog,
                    SelectionFormula = $"{{FCustomerCatalog.CustCode}}='{selectedCustomer.CustCode}'"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintStockList(bool combined)
        {
            if (selectedCustomer == null) return;
            try
            {
                var request = new ReportRequest
                {
                    ReportName = combined ? ReportNames.Company.CombCustStockList : ReportNames.Company.CustStockList,
                    SelectionFormula = $"{{VItems_StockReport.InActive}}=FALSE AND {{FCustomerCatalog.CustCode}}='{selectedCustomer.CustCode}'"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", ex.Message);
            }
        }

        private void ClosePage()
        {
            NavigationManager.NavigateTo("/company/dashboard");
        }
    }
}
