using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.FinishMovement
{
    public partial class FinishMovement : ComponentBase
    {
        [SupplyParameterFromQuery]
        public string ReturnUrl { get; set; }

        private string CurrentUserName = "System";
        private bool IsSaving = false;

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

        private List<FinishFromLocationModel> FromLocations = new();
        private FinishFromLocationModel SelectedFromLocation;

        private List<SFTargetStoreModel> TargetStores = new();
        private int SelectedTargetStoreId = 0;

        private List<SFTargetRackModel> TargetRacks = new();
        private string SelectedTargetRackNo = string.Empty;

        private List<SFTargetShelfModel> TargetShelves = new();
        private SFTargetShelfModel SelectedTargetShelf;

        private double TargetShelfQty = 0;
        private double MoveQty = 0;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                Articles = await MovementService.GetArticlesAsync();
                TargetStores = await MovementService.GetStoresAsync(CurrentUserName);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Initialization Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
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
            SelectedFromLocation = null;
            FromLocations.Clear();
            ResetTargetSelection();

            if (SelectedArticle == null)
            {
                await InvokeAsync(StateHasChanged);
                return;
            }

            try
            {
                FromLocations = await MovementService.GetFromLocationsAsync(SelectedArticle.ItemID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private void OnFromLocationChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int entryId) && entryId > 0)
            {
                SelectedFromLocation = FromLocations.FirstOrDefault(x => x.EntryID == entryId);
                MoveQty = SelectedFromLocation?.NetQty ?? 0;
            }
            else
            {
                SelectedFromLocation = null;
                MoveQty = 0;
            }
        }

        private async Task OnTargetStoreChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int storeId) && storeId > 0)
            {
                SelectedTargetStoreId = storeId;
                TargetRacks = await MovementService.GetRacksAsync(storeId);
            }
            else
            {
                SelectedTargetStoreId = 0;
                TargetRacks.Clear();
            }

            SelectedTargetRackNo = string.Empty;
            TargetShelves.Clear();
            SelectedTargetShelf = null;
            TargetShelfQty = 0;

            await InvokeAsync(StateHasChanged);
        }

        private async Task OnTargetRackChanged(ChangeEventArgs e)
        {
            string rackNo = e.Value?.ToString() ?? string.Empty;
            SelectedTargetRackNo = rackNo;

            if (!string.IsNullOrEmpty(rackNo) && SelectedTargetStoreId > 0)
            {
                TargetShelves = await MovementService.GetShelvesAsync(SelectedTargetStoreId, rackNo);
            }
            else
            {
                TargetShelves.Clear();
            }

            SelectedTargetShelf = null;
            TargetShelfQty = 0;

            await InvokeAsync(StateHasChanged);
        }

        private async Task OnTargetShelfChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int shelfId) && shelfId > 0)
            {
                SelectedTargetShelf = TargetShelves.FirstOrDefault(x => x.EntryID == shelfId);
                if (SelectedArticle != null && SelectedTargetShelf != null)
                {
                    TargetShelfQty = await MovementService.GetTargetShelfQtyAsync(SelectedArticle.ItemID, SelectedTargetShelf.EntryID);
                }
            }
            else
            {
                SelectedTargetShelf = null;
                TargetShelfQty = 0;
            }

            await InvokeAsync(StateHasChanged);
        }

        private void ResetTargetSelection()
        {
            SelectedTargetStoreId = 0;
            TargetRacks.Clear();
            SelectedTargetRackNo = string.Empty;
            TargetShelves.Clear();
            SelectedTargetShelf = null;
            TargetShelfQty = 0;
            MoveQty = 0;
        }

        private void ResetForm()
        {
            _selectedArticle = null;
            SelectedFromLocation = null;
            FromLocations.Clear();
            ResetTargetSelection();
            InvokeAsync(StateHasChanged);
        }

        private async Task SaveMovementAsync()
        {
            await SaveAsync(true);
        }

        private async Task SaveAsync(bool closeAfterSave)
        {
            if (SelectedArticle == null)
            {
                ShowWarning("Please select an article first.");
                return;
            }
            if (SelectedFromLocation == null)
            {
                ShowWarning("Please select a valid source location (Move From).");
                return;
            }
            if (SelectedTargetShelf == null)
            {
                ShowWarning("Please select a valid target store, rack, and shelf (Move To).");
                return;
            }
            if (MoveQty <= 0)
            {
                ShowWarning("Quantity to move must be greater than 0.");
                return;
            }
            if (MoveQty > SelectedFromLocation.NetQty)
            {
                ShowWarning($"Cannot move {MoveQty:N0}. Only {SelectedFromLocation.NetQty:N0} available in selected source location.");
                return;
            }

            IsSaving = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                var store = TargetStores.FirstOrDefault(s => s.EntryID == SelectedTargetStoreId);

                var req = new FinishMovementRequest
                {
                    ItemID = SelectedArticle.ItemID,
                    ItemName = SelectedArticle.ItemName,
                    SourceLocation = SelectedFromLocation,
                    TargetStoreRefId = SelectedTargetStoreId,
                    TargetStoreName = store?.StoreName ?? string.Empty,
                    TargetRackNo = SelectedTargetRackNo,
                    TargetShelfRefId = SelectedTargetShelf.EntryID,
                    TargetShelfNo = SelectedTargetShelf.ShelfNo,
                    MoveQty = MoveQty,
                    UserName = CurrentUserName
                };

                bool success = await MovementService.ExecuteMovementAsync(req);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Success",
                        Detail = $"Moved {MoveQty:N0} units of article {SelectedArticle.ItemID} successfully.",
                        Duration = 4000
                    });

                    if (closeAfterSave)
                    {
                        NavigateBack();
                    }
                    else
                    {
                        // Reset selection for new movement
                        _selectedArticle = null;
                        SelectedFromLocation = null;
                        FromLocations.Clear();
                        ResetTargetSelection();
                    }
                }
                else
                {
                    ShowError("Failed to save location movement record.");
                }
            }
            catch (Exception ex)
            {
                ShowError("Error saving movement: " + ex.Message);
            }
            finally
            {
                IsSaving = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private void ShowWarning(string message)
        {
            NotificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Warning,
                Summary = "Validation Warning",
                Detail = message,
                Duration = 4000
            });
        }

        private void ShowError(string message)
        {
            NotificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Error,
                Summary = "Error",
                Detail = message,
                Duration = 4000
            });
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
    }
}
