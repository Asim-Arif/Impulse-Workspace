using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.RMMovement
{
    public partial class RMMovement : ComponentBase
    {
        private bool IsLoadingInitialData = true;
        private bool IsSaving = false;

        private List<RMViewModel> Materials = new();
        private List<LocationwiseStatusViewModel> Locations = new();
        private List<StoreViewModel> Stores = new();
        private List<RackViewModel> Racks = new();
        private List<ShelfViewModel> Shelves = new();

        private RMViewModel _selectedMaterial;
        public RMViewModel SelectedMaterial
        {
            get => _selectedMaterial;
            set
            {
                if (_selectedMaterial != value)
                {
                    _selectedMaterial = value;
                    SelectedLocation = null;
                    QtyToMove = 0;
                    Locations.Clear();
                    _ = InvokeAsync(LoadLocationsAsync);
                }
            }
        }

        public LocationwiseStatusViewModel SelectedLocation { get; set; }

        private StoreViewModel _selectedStore;
        public StoreViewModel SelectedStore
        {
            get => _selectedStore;
            set
            {
                if (_selectedStore != value)
                {
                    _selectedStore = value;
                    SelectedRack = null;
                    SelectedShelf = null;
                    Racks.Clear();
                    Shelves.Clear();
                    _ = InvokeAsync(LoadRacksAsync);
                }
            }
        }

        private RackViewModel _selectedRack;
        public RackViewModel SelectedRack
        {
            get => _selectedRack;
            set
            {
                if (_selectedRack != value)
                {
                    _selectedRack = value;
                    SelectedShelf = null;
                    Shelves.Clear();
                    _ = InvokeAsync(LoadShelvesAsync);
                }
            }
        }

        private ShelfViewModel _selectedShelf;
        public ShelfViewModel SelectedShelf
        {
            get => _selectedShelf;
            set
            {
                if (_selectedShelf != value)
                {
                    _selectedShelf = value;
                    _ = InvokeAsync(LoadTargetShelfQtyAsync);
                }
            }
        }

        public double TargetShelfQty { get; set; }
        public double QtyToMove { get; set; }

        private string _userName;
        private string _machineName;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                _userName = authState.User.Identity?.Name ?? "System";
                _machineName = Environment.MachineName;

                // Load Initial Data
                Materials = await RMDataAccess.GetMaterialsAsync();
                
                int userId = 1;
                var idClaim = authState.User.FindFirst("UserId");
                if (idClaim != null && int.TryParse(idClaim.Value, out int uid))
                {
                    userId = uid;
                }
                Stores = await PlacementService.GetStoresForUserAsync(userId);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading data", ex.Message);
            }
            finally
            {
                IsLoadingInitialData = false;
            }
        }

        private async Task LoadLocationsAsync()
        {
            try
            {
                if (SelectedMaterial != null)
                {
                    Locations = await MovementService.GetLocationsForMaterialAsync(SelectedMaterial.RMID1);
                }
                else
                {
                    Locations.Clear();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading locations", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task LoadRacksAsync()
        {
            try
            {
                if (SelectedStore != null)
                {
                    Racks = await PlacementService.GetRacksAsync(SelectedStore.EntryID);
                }
                else
                {
                    Racks.Clear();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading racks", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task LoadShelvesAsync()
        {
            try
            {
                if (SelectedRack != null)
                {
                    Shelves = await PlacementService.GetShelvesAsync(SelectedRack.EntryID);
                }
                else
                {
                    Shelves.Clear();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading shelves", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task LoadTargetShelfQtyAsync()
        {
            try
            {
                if (SelectedShelf != null && SelectedMaterial != null)
                {
                    TargetShelfQty = await MovementService.GetShelfQuantityAsync(SelectedShelf.EntryID, SelectedMaterial.RMID1);
                }
                else
                {
                    TargetShelfQty = 0;
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading shelf quantity", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        // Typeahead Search Methods
        private async Task<IEnumerable<RMViewModel>> SearchMaterials(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Materials;
                
            return await Task.FromResult(Materials.Where(x => x.RMName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<LocationwiseStatusViewModel>> SearchLocations(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Locations;
                
            return await Task.FromResult(Locations.Where(x => x.DisplayText.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<StoreViewModel>> SearchStores(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Stores;
                
            return await Task.FromResult(Stores.Where(x => x.StoreName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<RackViewModel>> SearchRacks(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Racks;
                
            return await Task.FromResult(Racks.Where(x => x.RackNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<ShelfViewModel>> SearchShelves(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Shelves;
                
            return await Task.FromResult(Shelves.Where(x => x.ShelfNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private void ResetForm()
        {
            SelectedMaterial = null;
            SelectedStore = null;
            QtyToMove = 0;
        }

        private async Task SaveMovementAsync()
        {
            if (SelectedMaterial == null)
            {
                NotificationService.ShowError("Validation Error", "Please select a material.");
                return;
            }
            if (SelectedLocation == null)
            {
                NotificationService.ShowError("Validation Error", "Please select a source location.");
                return;
            }
            if (SelectedShelf == null)
            {
                NotificationService.ShowError("Validation Error", "Please select a target shelf.");
                return;
            }
            if (QtyToMove <= 0)
            {
                NotificationService.ShowError("Validation Error", "Quantity to move must be greater than zero.");
                return;
            }
            if (QtyToMove > SelectedLocation.AvailableQty)
            {
                NotificationService.ShowError("Validation Error", "Cannot move more than available quantity.");
                return;
            }

            bool confirm = await NotificationService.ShowQuestionNotification("Confirm Movement", 
                $"Are you sure you want to move {QtyToMove} of {SelectedMaterial.RMName} to {SelectedStore.StoreName} [{SelectedShelf.ShelfNo}]?");
            
            if (!confirm) return;

            IsSaving = true;
            try
            {
                var request = new RMMovementSaveRequest
                {
                    From_MLS_RefID = SelectedLocation.EntryID,
                    To_Shelf_RefID = SelectedShelf.EntryID,
                    QtyMoved = QtyToMove,
                    Rcvd_RefID = SelectedLocation.Rcvd_RefID,
                    BatchNo = SelectedLocation.BatchNo,
                    LotNo = SelectedLocation.LotNo
                };

                await MovementService.SaveMovementAsync(request, _userName, _machineName);
                
                NotificationService.ShowSuccess("Success", "Material moved successfully.");
                ResetForm();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }
    }
}
