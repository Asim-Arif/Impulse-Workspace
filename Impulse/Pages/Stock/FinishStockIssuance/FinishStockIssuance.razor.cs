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

namespace Impulse.Pages.Stock.FinishStockIssuance
{
    public partial class FinishStockIssuance : ComponentBase
    {
        private bool IsLoadingInitial = true;
        private bool IsSaving = false;
        private string CurrentUserName = "System";
        private string ReturnUrl;

        private FinishStockIssuanceMasterModel Master = new();
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

        private List<SFTargetStoreModel> AvailableStores = new();
        private int SelectedStoreId { get; set; } = 0;

        private List<LocationLotStockModel> AvailableLocations = new();
        private int SelectedLocationId { get; set; } = 0;

        private double AvailableStockQty { get; set; } = 0;
        private double LineQty { get; set; } = 0;
        private string LineRemarks { get; set; } = string.Empty;

        private List<FinishStockIssuanceItemModel> LineItems = new();

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

                Master.IssNo = await IssuanceService.GetNextIssuanceNoAsync(Master.DT);
                Customers = await IssuanceService.GetCustomersAsync();
                var firstCust = Customers.FirstOrDefault();
                if (firstCust != null)
                {
                    Master.CustCode = firstCust.CustCode;
                }

                Articles = await IssuanceService.GetArticlesAsync(Master.CustCode);
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
            Articles = await IssuanceService.GetArticlesAsync(Master.CustCode);
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

            try
            {
                AvailableStores = await IssuanceService.GetAvailableStoresAsync(SelectedArticle.ItemID);
                var firstStore = AvailableStores.FirstOrDefault();
                if (firstStore != null)
                {
                    SelectedStoreId = firstStore.EntryID;
                    await LoadLocationsAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading stores", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task OnStoreChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int storeId))
            {
                SelectedStoreId = storeId;
                AvailableLocations.Clear();
                SelectedLocationId = 0;
                AvailableStockQty = 0;
                LineQty = 0;

                if (SelectedStoreId > 0 && SelectedArticle != null)
                {
                    await LoadLocationsAsync();
                }
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task LoadLocationsAsync()
        {
            if (SelectedArticle == null || SelectedStoreId == 0) return;

            AvailableLocations = await IssuanceService.GetAvailableLocationsAsync(SelectedArticle.ItemID, SelectedStoreId, Master.CustCode);
            var firstLoc = AvailableLocations.FirstOrDefault();
            if (firstLoc != null)
            {
                SelectedLocationId = firstLoc.EntryID;
                AvailableStockQty = firstLoc.NetQty;
                LineQty = AvailableStockQty;
            }
            else
            {
                SelectedLocationId = 0;
                AvailableStockQty = 0;
                LineQty = 0;
            }
            await InvokeAsync(StateHasChanged);
        }

        private void OnLocationChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int locId))
            {
                SelectedLocationId = locId;
                var loc = AvailableLocations.FirstOrDefault(x => x.EntryID == SelectedLocationId);
                if (loc != null)
                {
                    AvailableStockQty = loc.NetQty;
                    LineQty = AvailableStockQty;
                }
                else
                {
                    AvailableStockQty = 0;
                    LineQty = 0;
                }
            }
            StateHasChanged();
        }

        private void AddLineItem()
        {
            if (SelectedArticle == null)
            {
                NotificationService.ShowWarning("Validation", "Please select an Article.");
                return;
            }

            if (SelectedLocationId == 0)
            {
                NotificationService.ShowWarning("Validation", "Please select a Location.");
                return;
            }

            if (LineQty <= 0)
            {
                NotificationService.ShowWarning("Validation", "Quantity to issue must be greater than 0.");
                return;
            }

            if (LineQty > AvailableStockQty)
            {
                NotificationService.ShowWarning("Validation", $"Only {AvailableStockQty:N0} is available at selected location.");
                return;
            }

            var loc = AvailableLocations.FirstOrDefault(x => x.EntryID == SelectedLocationId);
            if (loc == null) return;

            string locationDisplay = $"{loc.StoreName} [{loc.RackNo}] [{loc.ShelfNo}]";

            var newItem = new FinishStockIssuanceItemModel
            {
                LineNo = LineItems.Count + 1,
                ItemID = SelectedArticle.ItemID,
                ItemName = SelectedArticle.ItemName,
                Unit = SelectedUnit,
                Store_RefID = SelectedStoreId,
                StoreName = loc.StoreName,
                Shelf_RefID = loc.Shelf_RefID,
                LocationDisplay = locationDisplay,
                RISD_EntryID = loc.EntryID,
                LotNo = loc.LotNo,
                BatchNo = loc.Batch_No,
                AvailableQty = AvailableStockQty,
                Qty = LineQty,
                Remarks = LineRemarks
            };

            LineItems.Add(newItem);
            NotificationService.ShowSuccess("Line Added", $"Added {newItem.ItemName} ({newItem.Qty:N0} {newItem.Unit})");

            ClearSelectedArticleAndLine();
            StateHasChanged();
        }

        private void RemoveLineItem(FinishStockIssuanceItemModel item)
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
            if (string.IsNullOrWhiteSpace(Master.Department))
            {
                NotificationService.ShowWarning("Validation", "Please enter Department.");
                return;
            }

            if (string.IsNullOrWhiteSpace(Master.AttnPerson))
            {
                NotificationService.ShowWarning("Validation", "Please enter Attention Person.");
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
                bool success = await IssuanceService.SaveIssuanceAsync(Master, LineItems);
                if (success)
                {
                    NotificationService.ShowSuccess("Saved Successfully", $"Issuance #{Master.IssNo} saved successfully.");
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
                    NotificationService.ShowError("Error", "Failed to save issuance.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Saving Issuance", ex.Message);
            }
            finally
            {
                IsSaving = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private void ResetLineEntry()
        {
            AvailableStores.Clear();
            SelectedStoreId = 0;
            AvailableLocations.Clear();
            SelectedLocationId = 0;
            AvailableStockQty = 0;
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
            Master = new FinishStockIssuanceMasterModel
            {
                DT = DateTime.Today,
                UserName = CurrentUserName,
                CustCode = Customers.FirstOrDefault()?.CustCode ?? "Stock"
            };

            Master.IssNo = await IssuanceService.GetNextIssuanceNoAsync(Master.DT);
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
