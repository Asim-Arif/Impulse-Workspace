using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.WebUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.FinishStockReceiving
{
    public partial class FinishStockReceiving : ComponentBase
    {
        private bool IsLoadingInitial = true;
        private bool IsSaving = false;
        private string CurrentUserName = "System";
        private string ReturnUrl;

        private FinishStockReceivingMasterModel Master = new();
        private List<CustomerLookupModel> Customers = new();

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
                    _ = OnArticleChanged();
                }
            }
        }

        private string SelectedUnit = string.Empty;

        private List<SFTargetStoreModel> Stores = new();
        private int SelectedStoreId { get; set; } = 0;

        private List<SFTargetShelfModel> Shelves = new();
        private int SelectedShelfId { get; set; } = 0;

        private string LotNoManual { get; set; } = string.Empty;
        private string BatchNoManual { get; set; } = string.Empty;
        private string MillCertNo { get; set; } = string.Empty;
        private double LineQty { get; set; } = 0;
        private string LineRemarks { get; set; } = string.Empty;

        private List<FinishStockReceivingItemModel> LineItems = new();

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
                Master.UserName = CurrentUserName;
                Master.RcvdBy = CurrentUserName;

                Master.RcvNo = await ReceivingService.GetNextReceivingNoAsync(Master.DT);
                Customers = await ReceivingService.GetCustomersAsync();
                var firstCust = Customers.FirstOrDefault();
                if (firstCust != null)
                {
                    Master.CustCode = firstCust.CustCode;
                }

                Articles = await ReceivingService.GetArticlesAsync(Master.CustCode);
                Stores = await ReceivingService.GetStoresAsync(CurrentUserName);
                var firstStore = Stores.FirstOrDefault();
                if (firstStore != null)
                {
                    SelectedStoreId = firstStore.EntryID;
                    Shelves = await ReceivingService.GetShelvesAsync(SelectedStoreId);
                    var firstShelf = Shelves.FirstOrDefault();
                    if (firstShelf != null)
                    {
                        SelectedShelfId = firstShelf.EntryID;
                    }
                }

                LotNoManual = await ReceivingService.GetNextLotNoAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
            finally
            {
                IsLoadingInitial = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task<IEnumerable<ItemViewModel>> SearchArticles(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Articles;

            return await Task.FromResult(Articles.Where(x => x.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnCustomerChanged(ChangeEventArgs e)
        {
            Master.CustCode = e.Value?.ToString() ?? "Stock";
            ClearSelectedArticleAndLine();
            Articles = await ReceivingService.GetArticlesAsync(Master.CustCode);
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnArticleChanged()
        {
            ResetLineEntry();

            if (SelectedArticle == null)
            {
                SelectedUnit = string.Empty;
                await InvokeAsync(StateHasChanged);
                return;
            }

            SelectedUnit = SelectedArticle.Unit;
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnStoreChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int storeId))
            {
                SelectedStoreId = storeId;
                Shelves.Clear();
                SelectedShelfId = 0;

                if (SelectedStoreId > 0)
                {
                    Shelves = await ReceivingService.GetShelvesAsync(SelectedStoreId);
                    var firstShelf = Shelves.FirstOrDefault();
                    if (firstShelf != null)
                    {
                        SelectedShelfId = firstShelf.EntryID;
                    }
                }
            }
            await InvokeAsync(StateHasChanged);
        }

        private void OnShelfChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int shelfId))
            {
                SelectedShelfId = shelfId;
            }
            StateHasChanged();
        }

        private async Task AddLineItem()
        {
            if (SelectedArticle == null)
            {
                NotificationService.ShowWarning("Validation", "Please select an Article.");
                return;
            }

            if (SelectedShelfId == 0)
            {
                NotificationService.ShowWarning("Validation", "Please select a Target Location.");
                return;
            }

            if (string.IsNullOrWhiteSpace(LotNoManual))
            {
                NotificationService.ShowWarning("Validation", "Please enter Lot No.");
                return;
            }

            if (LineQty <= 0)
            {
                NotificationService.ShowWarning("Validation", "Quantity to receive must be greater than 0.");
                return;
            }

            // Check if LotNo is already in current line items
            if (LineItems.Any(x => x.LotNo_Manual.Equals(LotNoManual, StringComparison.OrdinalIgnoreCase)))
            {
                NotificationService.ShowWarning("Validation", $"Lot No. '{LotNoManual}' is already added in this entry.");
                return;
            }

            // Check if LotNo exists in DB
            bool existsInDb = await ReceivingService.IsLotNoExistsAsync(LotNoManual);
            if (existsInDb)
            {
                NotificationService.ShowWarning("Validation", $"Lot No. '{LotNoManual}' already exists in database.");
                return;
            }

            var shelf = Shelves.FirstOrDefault(x => x.EntryID == SelectedShelfId);
            if (shelf == null) return;

            string locationDisplay = $"{shelf.StoreName} [{shelf.RackNo}] [{shelf.ShelfNo}]";

            var newItem = new FinishStockReceivingItemModel
            {
                LineNo = LineItems.Count + 1,
                ItemID = SelectedArticle.ItemID,
                ItemName = SelectedArticle.ItemName,
                Unit = SelectedUnit,
                Store_RefID = SelectedStoreId,
                StoreName = shelf.StoreName,
                Shelf_RefID = SelectedShelfId,
                LocationDisplay = locationDisplay,
                LotNo_Manual = LotNoManual,
                BatchNo_Manual = BatchNoManual,
                Mill_Certificate_No = MillCertNo,
                Qty = LineQty,
                Remarks = LineRemarks
            };

            LineItems.Add(newItem);
            NotificationService.ShowSuccess("Line Added", $"Added {newItem.ItemName} ({newItem.Qty:N0} {newItem.Unit})");

            // Auto increment Lot No for next item
            if (long.TryParse(LotNoManual, out long currentLot))
            {
                LotNoManual = (currentLot + 1).ToString();
            }

            ClearSelectedArticleAndLine();
            await InvokeAsync(StateHasChanged);
        }

        private void RemoveLineItem(FinishStockReceivingItemModel item)
        {
            LineItems.Remove(item);
            int seq = 1;
            foreach (var l in LineItems)
            {
                l.LineNo = seq++;
            }
            StateHasChanged();
        }

        private async Task SaveAsync(bool closeAfterSave)
        {
            if (string.IsNullOrWhiteSpace(Master.RcvdFrom))
            {
                NotificationService.ShowWarning("Validation", "Please enter Received From.");
                return;
            }

            if (!LineItems.Any())
            {
                NotificationService.ShowWarning("Validation", "Please add at least one line item.");
                return;
            }

            IsSaving = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                bool success = await ReceivingService.SaveReceivingAsync(Master, LineItems);
                if (success)
                {
                    NotificationService.ShowSuccess("Saved Successfully", $"Receiving #{Master.RcvNo} saved successfully.");
                    if (closeAfterSave)
                    {
                        NavigateBack();
                    }
                    else
                    {
                        await ResetFormAsync();
                    }
                }
                else
                {
                    NotificationService.ShowError("Error", "Failed to save receiving.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Saving Receiving", ex.Message);
            }
            finally
            {
                IsSaving = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private void ResetLineEntry()
        {
            BatchNoManual = string.Empty;
            MillCertNo = string.Empty;
            LineQty = 0;
            LineRemarks = string.Empty;
        }

        private void ClearSelectedArticleAndLine()
        {
            _selectedArticle = null;
            SelectedUnit = string.Empty;
            ResetLineEntry();
        }

        private async Task ResetFormAsync()
        {
            Master = new FinishStockReceivingMasterModel
            {
                DT = DateTime.Today,
                RcvdBy = CurrentUserName,
                UserName = CurrentUserName,
                CustCode = Customers.FirstOrDefault()?.CustCode ?? "Stock"
            };

            Master.RcvNo = await ReceivingService.GetNextReceivingNoAsync(Master.DT);
            LotNoManual = await ReceivingService.GetNextLotNoAsync();
            LineItems.Clear();
            ClearSelectedArticleAndLine();
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

        private void Cancel()
        {
            NavigateBack();
        }
    }
}
