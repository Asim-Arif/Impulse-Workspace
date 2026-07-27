using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.WebUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.SFMovement
{
    public partial class SFMovement : ComponentBase
    {
        private bool IsLoadingInitialData = true;
        private bool IsSaving = false;
        private string CurrentUserName = "System";
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
                    _ = OnArticleChanged();
                }
            }
        }

        private List<ItemProcessViewModel> Processes = new();
        private int SelectedProcessId { get; set; }

        private List<SFTargetStoreModel> FromStores = new();
        private int SelectedFromStoreId { get; set; }

        private List<SFFromLocationModel> FromLocations = new();
        private int SelectedFromLocationId { get; set; }
        private int AvailableQty { get; set; }
        private string LocationRemarks { get; set; } = string.Empty;

        private List<SFTargetStoreModel> TargetStores = new();
        private int SelectedTargetStoreId { get; set; }

        private List<SFTargetRackModel> TargetRacks = new();
        private string SelectedTargetRackNo { get; set; } = string.Empty;

        private List<SFTargetShelfModel> TargetShelves = new();
        private int SelectedTargetShelfId { get; set; }
        private int TargetShelfQty { get; set; }

        private int QtyToMove { get; set; }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var uri = NavManager.ToAbsoluteUri(NavManager.Uri);
                if (QueryHelpers.ParseQuery(uri.Query).TryGetValue("returnUrl", out var returnUrlStr))
                {
                    ReturnUrl = returnUrlStr;
                }

                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                Articles = await MovementService.GetArticlesAsync();
                TargetStores = await MovementService.GetTargetStoresAsync(CurrentUserName);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
            finally
            {
                IsLoadingInitialData = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task<IEnumerable<ItemViewModel>> SearchArticles(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Articles;

            return await Task.FromResult(Articles.Where(x => x.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnArticleChanged()
        {
            ResetFromSelection();
            Processes.Clear();
            SelectedProcessId = 0;

            if (SelectedArticle == null)
            {
                await InvokeAsync(StateHasChanged);
                return;
            }

            try
            {
                Processes = await MovementService.GetProcessesByArticleAsync(SelectedArticle.ItemID);
                var firstProc = Processes.FirstOrDefault();
                if (firstProc != null)
                {
                    SelectedProcessId = firstProc.ProcessID;
                    await LoadFromStoresAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading processes", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task OnProcessChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int procId))
            {
                SelectedProcessId = procId;
                ResetFromSelection();
                if (SelectedProcessId > 0 && SelectedArticle != null)
                {
                    await LoadFromStoresAsync();
                }
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task LoadFromStoresAsync()
        {
            if (SelectedArticle == null || SelectedProcessId == 0) return;

            FromStores = await MovementService.GetFromStoresAsync(SelectedArticle.ItemID, SelectedProcessId);
            var firstStore = FromStores.FirstOrDefault();
            if (firstStore != null)
            {
                SelectedFromStoreId = firstStore.EntryID;
                await LoadFromLocationsAsync();
            }
            else
            {
                SelectedFromStoreId = 0;
                FromLocations.Clear();
                SelectedFromLocationId = 0;
                AvailableQty = 0;
                LocationRemarks = string.Empty;
                QtyToMove = 0;
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnFromStoreChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int storeId))
            {
                SelectedFromStoreId = storeId;
                FromLocations.Clear();
                SelectedFromLocationId = 0;
                AvailableQty = 0;
                LocationRemarks = string.Empty;
                QtyToMove = 0;

                if (SelectedFromStoreId > 0 && SelectedArticle != null && SelectedProcessId > 0)
                {
                    await LoadFromLocationsAsync();
                }
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task LoadFromLocationsAsync()
        {
            FromLocations = await MovementService.GetFromLocationsAsync(SelectedArticle.ItemID, SelectedProcessId, SelectedFromStoreId);
            var firstLoc = FromLocations.FirstOrDefault();
            if (firstLoc != null)
            {
                SelectedFromLocationId = firstLoc.EntryID;
                AvailableQty = firstLoc.AvailableQty;
                LocationRemarks = firstLoc.Remarks;
                QtyToMove = AvailableQty;
            }
            else
            {
                SelectedFromLocationId = 0;
                AvailableQty = 0;
                LocationRemarks = string.Empty;
                QtyToMove = 0;
            }
            await InvokeAsync(StateHasChanged);
        }

        private void OnFromLocationChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int locId))
            {
                SelectedFromLocationId = locId;
                var selected = FromLocations.FirstOrDefault(x => x.EntryID == SelectedFromLocationId);
                if (selected != null)
                {
                    AvailableQty = selected.AvailableQty;
                    LocationRemarks = selected.Remarks;
                    QtyToMove = AvailableQty;
                }
                else
                {
                    AvailableQty = 0;
                    LocationRemarks = string.Empty;
                    QtyToMove = 0;
                }
            }
            StateHasChanged();
        }

        private async Task OnTargetStoreChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int storeId))
            {
                SelectedTargetStoreId = storeId;
                TargetRacks.Clear();
                SelectedTargetRackNo = string.Empty;
                TargetShelves.Clear();
                SelectedTargetShelfId = 0;
                TargetShelfQty = 0;

                if (SelectedTargetStoreId > 0)
                {
                    TargetRacks = await MovementService.GetTargetRacksAsync(SelectedTargetStoreId);
                    var firstRack = TargetRacks.FirstOrDefault();
                    if (firstRack != null)
                    {
                        SelectedTargetRackNo = firstRack.RackNo;
                        await LoadTargetShelvesAsync();
                    }
                }
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnTargetRackChanged(ChangeEventArgs e)
        {
            SelectedTargetRackNo = e.Value?.ToString() ?? string.Empty;
            TargetShelves.Clear();
            SelectedTargetShelfId = 0;
            TargetShelfQty = 0;

            if (!string.IsNullOrEmpty(SelectedTargetRackNo) && SelectedTargetStoreId > 0)
            {
                await LoadTargetShelvesAsync();
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task LoadTargetShelvesAsync()
        {
            TargetShelves = await MovementService.GetTargetShelvesAsync(SelectedTargetStoreId, SelectedTargetRackNo);
            var firstShelf = TargetShelves.FirstOrDefault();
            if (firstShelf != null)
            {
                SelectedTargetShelfId = firstShelf.EntryID;
                await UpdateTargetShelfQtyAsync();
            }
            else
            {
                SelectedTargetShelfId = 0;
                TargetShelfQty = 0;
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnTargetShelfChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int shelfId))
            {
                SelectedTargetShelfId = shelfId;
                await UpdateTargetShelfQtyAsync();
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task UpdateTargetShelfQtyAsync()
        {
            if (SelectedTargetShelfId > 0 && SelectedArticle != null && SelectedProcessId > 0)
            {
                TargetShelfQty = await MovementService.GetTargetLocationQtyAsync(SelectedArticle.ItemID, SelectedProcessId, SelectedTargetShelfId);
            }
            else
            {
                TargetShelfQty = 0;
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task SaveMovementAsync()
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
            if (SelectedFromLocationId == 0)
            {
                NotificationService.ShowWarning("Validation", "Please select From Location.");
                return;
            }
            if (SelectedTargetShelfId == 0)
            {
                NotificationService.ShowWarning("Validation", "Please select Target Location.");
                return;
            }
            if (AvailableQty <= 0)
            {
                NotificationService.ShowWarning("Validation", "No available quantity to move.");
                return;
            }

            IsSaving = true;

            try
            {
                var targetShelf = TargetShelves.FirstOrDefault(x => x.EntryID == SelectedTargetShelfId);
                string targetLocationDisplay = targetShelf != null ? $"{targetShelf.RackNo} {{{targetShelf.ShelfNo}}}" : string.Empty;

                var request = new SFMovementRequest
                {
                    SOO_EntryID = SelectedFromLocationId,
                    ItemID = SelectedArticle.ItemID,
                    ProcessID = SelectedProcessId,
                    TargetShelfID = SelectedTargetShelfId,
                    TargetLocationDisplay = targetLocationDisplay,
                    QtyToMove = AvailableQty,
                    UserName = CurrentUserName
                };

                bool success = await MovementService.TransferLocationAsync(request);
                if (success)
                {
                    NotificationService.ShowSuccess("Success", "Material location updated successfully.");
                    if (!string.IsNullOrWhiteSpace(ReturnUrl))
                    {
                        NavigateBack();
                    }
                    else
                    {
                        ResetFromSelection();
                        await RefreshLocationsAsync();
                    }
                }
                else
                {
                    NotificationService.ShowError("Error", "Failed to update location.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task RefreshLocationsAsync()
        {
            if (SelectedArticle != null && SelectedProcessId > 0)
            {
                await LoadFromStoresAsync();
                await UpdateTargetShelfQtyAsync();
            }
        }

        private void ResetFromSelection()
        {
            FromStores.Clear();
            SelectedFromStoreId = 0;
            FromLocations.Clear();
            SelectedFromLocationId = 0;
            AvailableQty = 0;
            LocationRemarks = string.Empty;
            QtyToMove = 0;
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavManager.NavigateTo("/stock");
            }
        }

        private void ResetForm()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigateBack();
                return;
            }

            _selectedArticle = null;
            ResetFromSelection();
            Processes.Clear();
            SelectedProcessId = 0;

            TargetStores.Clear();
            SelectedTargetStoreId = 0;
            TargetRacks.Clear();
            SelectedTargetRackNo = string.Empty;
            TargetShelves.Clear();
            SelectedTargetShelfId = 0;
            TargetShelfQty = 0;

            _ = OnInitializedAsync();
        }
    }
}
