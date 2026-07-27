using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.WebUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.SemiFinishOpenReceiving
{
    public partial class SemiFinishOpenReceiving : ComponentBase
    {
        [Inject] private NavigationManager NavigationManager { get; set; }

        private bool IsLoadingInitialData = true;
        private bool IsLoadingGrid = false;
        private bool IsSaving = false;

        private string CurrentUserName = "System";
        private string MachineName = "WebClient";
        private string ReturnUrl;

        private List<ItemViewModel> Articles = new();
        private ItemViewModel _selectedArticle;
        public ItemViewModel SelectedArticle
        {
            get => _selectedArticle;
            set
            {
                if (_selectedArticle != value)
                {
                    _selectedArticle = value;
                    _ = LoadArticleDetailsAsync();
                }
            }
        }

        private List<ItemProcessViewModel> Processes = new();
        private int SelectedProcessId { get; set; }

        private List<StoreLookupModel> Stores = new();
        private int SelectedStoreId { get; set; }

        private List<StoreShelfLookupModel> Shelves = new();
        private int SelectedShelfId { get; set; }

        private int Qty { get; set; } = 1;
        private string Remarks { get; set; } = string.Empty;
        private string LotNo { get; set; } = string.Empty;
        private string BatchNo { get; set; } = string.Empty;
        private string MillNo { get; set; } = string.Empty;

        private List<SemiFinishOpeningModel> Openings = new();

        // Edit Remarks Modal
        private bool ShowEditModal = false;
        private SemiFinishOpeningModel EditingItem;
        private string ModalRemarksText = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var uri = NavigationManager.ToAbsoluteUri(NavigationManager.Uri);
                if (QueryHelpers.ParseQuery(uri.Query).TryGetValue("returnUrl", out var returnUrlStr))
                {
                    ReturnUrl = returnUrlStr;
                }

                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                Articles = await Service.GetArticlesAsync();
                Stores = await Service.GetUserStoresAsync(CurrentUserName);
                LotNo = await Service.GetNextLotNoAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
            finally
            {
                IsLoadingInitialData = false;
            }
        }

        private async Task<IEnumerable<ItemViewModel>> SearchArticles(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Articles;

            return await Task.FromResult(Articles.Where(x => x.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task LoadArticleDetailsAsync()
        {
            if (SelectedArticle == null)
            {
                Processes.Clear();
                SelectedProcessId = 0;
                Openings.Clear();
                return;
            }

            IsLoadingGrid = true;
            StateHasChanged();

            try
            {
                Processes = await Service.GetProcessesByArticleAsync(SelectedArticle.ItemID);
                SelectedProcessId = Processes.FirstOrDefault()?.ProcessID ?? 0;

                Openings = await Service.GetOpeningsByArticleAsync(SelectedArticle.ItemID);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading details", ex.Message);
            }
            finally
            {
                IsLoadingGrid = false;
                StateHasChanged();
            }
        }

        private async Task OnStoreChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int storeId))
            {
                SelectedStoreId = storeId;
                SelectedShelfId = 0;
                if (SelectedStoreId > 0)
                {
                    Shelves = await Service.GetStoreShelvesAsync(SelectedStoreId);
                    SelectedShelfId = Shelves.FirstOrDefault()?.EntryID ?? 0;
                }
                else
                {
                    Shelves.Clear();
                }
            }
        }

        private async Task SaveOpeningAsync()
        {
            if (SelectedArticle == null)
            {
                NotificationService.ShowWarning("Validation", "Please select an Article.");
                return;
            }
            if (SelectedProcessId == 0)
            {
                NotificationService.ShowWarning("Validation", "Please select a Process.");
                return;
            }
            if (Qty <= 0)
            {
                NotificationService.ShowWarning("Validation", "Please enter a valid Quantity.");
                return;
            }
            if (SelectedShelfId == 0)
            {
                NotificationService.ShowWarning("Validation", "Please select a Location.");
                return;
            }
            if (string.IsNullOrWhiteSpace(BatchNo))
            {
                NotificationService.ShowWarning("Validation", "Please enter Batch No.");
                return;
            }

            IsSaving = true;

            try
            {
                // Check if Batch No exists
                bool batchExists = await Service.CheckBatchNoExistsAsync(BatchNo);
                if (batchExists)
                {
                    NotificationService.ShowWarning("Batch Warning", $"Batch No '{BatchNo}' already exists.");
                }

                var selectedShelf = Shelves.FirstOrDefault(x => x.EntryID == SelectedShelfId);

                var request = new CreateSemiFinishOpeningRequest
                {
                    ItemID = SelectedArticle.ItemID,
                    ProcessID = SelectedProcessId,
                    Qty = Qty,
                    Shelf_RefID = SelectedShelfId,
                    Location = selectedShelf != null ? $"{selectedShelf.RackNo} {{{selectedShelf.ShelfNo}}}" : string.Empty,
                    Remarks = Remarks,
                    LotNo_Manual = LotNo,
                    BatchNo_Manual = BatchNo,
                    Mill_Certificate_No = MillNo,
                    UserName = CurrentUserName,
                    MachineName = MachineName
                };

                int newEntryId = await Service.AddOpeningAsync(request);
                NotificationService.ShowSuccess("Success", $"Semi Finish Opening entry #{newEntryId} saved successfully.");

                // Print SF Opening Slip report
                var printReq = new ReportRequest
                {
                    ReportName = ReportNames.Stock.SFOpeningSlip,
                    SelectionFormula = $"{{StockOrderOpening.EntryID}}={newEntryId}"
                };
                await ReportNavigation.PrintReportAsync(printReq);

                // Reset form fields
                Qty = 1;
                Remarks = string.Empty;
                BatchNo = string.Empty;
                MillNo = string.Empty;
                LotNo = await Service.GetNextLotNoAsync();

                // Refresh grid
                Openings = await Service.GetOpeningsByArticleAsync(SelectedArticle.ItemID);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        private async Task PrintLedgerAsync()
        {
            if (SelectedArticle == null) return;

            var request = new ReportRequest
            {
                ReportName = ReportNames.Stock.SFLedger,
                SelectionFormula = $"{{VStockOrderOpening.ItemID}}='{SelectedArticle.ItemID}'"
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintTrayCard(SemiFinishOpeningModel item)
        {
            if (item == null) return;

            var request = new ReportRequest
            {
                ReportName = ReportNames.Stock.SFTrayCard,
                SelectionFormula = $"{{VStockOrderOpening.EntryID}}={item.EntryID}"
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private void OpenEditRemarksModal(SemiFinishOpeningModel item)
        {
            EditingItem = item;
            ModalRemarksText = item?.Remarks ?? string.Empty;
            ShowEditModal = true;
        }

        private void CloseEditModal()
        {
            ShowEditModal = false;
            EditingItem = null;
            ModalRemarksText = string.Empty;
        }

        private async Task SaveModalRemarks()
        {
            if (EditingItem == null) return;

            try
            {
                await Service.UpdateRemarksAsync(EditingItem.EntryID, ModalRemarksText);
                EditingItem.Remarks = ModalRemarksText;
                NotificationService.ShowSuccess("Updated", "Remarks updated successfully.");
                CloseEditModal();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Update Error", ex.Message);
            }
        }

        private async Task ConfirmDelete(SemiFinishOpeningModel item)
        {
            if (item == null) return;

            try
            {
                bool deleted = await Service.DeleteOpeningAsync(item.EntryID);
                if (!deleted)
                {
                    NotificationService.ShowWarning("Delete Restricted", "Can't Delete. Issuance has already been made against this entry.");
                    return;
                }

                NotificationService.ShowSuccess("Deleted", "Entry deleted successfully.");
                if (SelectedArticle != null)
                {
                    Openings = await Service.GetOpeningsByArticleAsync(SelectedArticle.ItemID);
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Delete Error", ex.Message);
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/stock");
            }
        }

        // Context Menu Event Handlers
        private async Task OnContextPrintTrayCard(ItemClickEventArgs e)
        {
            if (e.Data is SemiFinishOpeningModel item)
            {
                await PrintTrayCard(item);
            }
        }

        private void OnContextEditRemarks(ItemClickEventArgs e)
        {
            if (e.Data is SemiFinishOpeningModel item)
            {
                OpenEditRemarksModal(item);
            }
        }

        private async Task OnContextDelete(ItemClickEventArgs e)
        {
            if (e.Data is SemiFinishOpeningModel item)
            {
                await ConfirmDelete(item);
            }
        }
    }
}
