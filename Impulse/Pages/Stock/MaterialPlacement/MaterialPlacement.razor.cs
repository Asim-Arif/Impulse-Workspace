using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;

namespace Impulse.Pages.Stock.MaterialPlacement
{
    public partial class MaterialPlacement : ComponentBase
    {
        [Inject] private IConfiguration Configuration { get; set; }

        public class DropdownVendor { public string VendID { get; set; } public string VendorName { get; set; } }
        public class DropdownMaterial { public string MaterialID { get; set; } public string MaterialName { get; set; } }

        private List<MaterialPlacementListViewModel> PendingPlacements;
        private DropdownVendor SelectedVendor;
        private DropdownMaterial SelectedMaterial;
        
        private bool IsLoading = false;

        // Modal State
        private bool IsModalOpen = false;
        private MaterialPlacementListViewModel SelectedPlacementItem;
        
        public bool ShowPicture { get; set; }
        public string SelectedMaterialId { get; set; }

        private void SelectRowMaterial(string materialId)
        {
            SelectedMaterialId = materialId;
        }
        private StoreViewModel _selectedStore;
        private StoreViewModel SelectedStore
        {
            get => _selectedStore;
            set
            {
                if (_selectedStore != value)
                {
                    _selectedStore = value;
                    SelectedRack = null;
                    SelectedShelf = null;
                    AvailableRacks.Clear();
                    AvailableShelves.Clear();
                }
            }
        }

        private RackViewModel _selectedRack;
        private RackViewModel SelectedRack
        {
            get => _selectedRack;
            set
            {
                if (_selectedRack != value)
                {
                    _selectedRack = value;
                    SelectedShelf = null;
                    AvailableShelves.Clear();
                }
            }
        }

        private ShelfViewModel SelectedShelf;
        private decimal PlacementQty;
        private string BatchNo;
        private string LotNo;

        private List<StoreViewModel> UserStores = new List<StoreViewModel>();
        private List<RackViewModel> AvailableRacks = new List<RackViewModel>();
        private List<ShelfViewModel> AvailableShelves = new List<ShelfViewModel>();

        private int CurrentUserId = 0;
        private string CurrentUserName = "System";

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
            if (authState.User.Identity.IsAuthenticated)
            {
                CurrentUserName = authState.User.Identity.Name;
                var idClaim = authState.User.FindFirst("UserId");
                if (idClaim != null && int.TryParse(idClaim.Value, out int uid))
                {
                    CurrentUserId = uid;
                }
            }

            // Load Initial Grid
            await LoadGridDataAsync();
        }

        private async Task LoadGridDataAsync()
        {
            IsLoading = true;
            StateHasChanged();
            
            try
            {
                string vId = SelectedVendor?.VendID;
                string mId = SelectedMaterial?.MaterialID;
                PendingPlacements = await MaterialPlacementService.GetPendingPlacementsAsync(vId, mId);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", "Failed to load placements: " + ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task<IEnumerable<DropdownVendor>> SearchVendors(string searchText)
        {
            using (IDbConnection db = new SqlConnection(Configuration.GetConnectionString("DefaultConnection")))
            {
                string sql = "SELECT AccNo AS VendID, AccTitle AS VendorName FROM VVenders WHERE AccTitle LIKE @Search ORDER BY AccTitle";
                return await db.QueryAsync<DropdownVendor>(sql, new { Search = $"%{searchText}%" });
            }
        }

        private async Task<IEnumerable<DropdownMaterial>> SearchMaterials(string searchText)
        {
            using (IDbConnection db = new SqlConnection(Configuration.GetConnectionString("DefaultConnection")))
            {
                string sql = "SELECT RMID1 AS MaterialID, RMName AS MaterialName FROM RM WHERE RMName LIKE @Search ORDER BY RMName";
                return await db.QueryAsync<DropdownMaterial>(sql, new { Search = $"%{searchText}%" });
            }
        }

        private async Task OpenPlacementModal(MaterialPlacementListViewModel item)
        {
            SelectedPlacementItem = item;
            SelectedStore = null;
            SelectedRack = null;
            SelectedShelf = null;
            PlacementQty = item.Balance;
            BatchNo = "";
            LotNo = "";
            
            UserStores = await MaterialPlacementService.GetStoresForUserAsync(CurrentUserId);
            AvailableRacks.Clear();
            AvailableShelves.Clear();

            IsModalOpen = true;
        }

        private void CloseModal()
        {
            IsModalOpen = false;
            SelectedPlacementItem = null;
        }

        private async Task<IEnumerable<StoreViewModel>> SearchStores(string searchText)
        {
            return await Task.FromResult(UserStores.Where(x => x.StoreName.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task<IEnumerable<RackViewModel>> SearchRacks(string searchText)
        {
            if (SelectedStore != null && !AvailableRacks.Any())
            {
                AvailableRacks = await MaterialPlacementService.GetRacksAsync(SelectedStore.EntryID);
            }
            return AvailableRacks.Where(x => x.RackNo.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task<IEnumerable<ShelfViewModel>> SearchShelves(string searchText)
        {
            if (SelectedRack != null && !AvailableShelves.Any())
            {
                AvailableShelves = await MaterialPlacementService.GetShelvesAsync(SelectedRack.EntryID);
            }
            return AvailableShelves.Where(x => x.ShelfNo.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task SavePlacementAsync()
        {
            if (PlacementQty <= 0 || PlacementQty > SelectedPlacementItem.Balance)
            {
                NotificationService.ShowWarning("Warning", "Invalid placement quantity.");
                return;
            }
            if (SelectedStore == null || SelectedRack == null || SelectedShelf == null)
            {
                NotificationService.ShowWarning("Warning", "Please select a Store, Rack, and Shelf.");
                return;
            }

            try
            {
                var request = new PlacementSaveRequest
                {
                    Rcvd_RefID = SelectedPlacementItem.EntryID, // Tag from VB6
                    Shelf_RefID = SelectedShelf.EntryID,
                    QtyPlaced = PlacementQty,
                    BatchNo = BatchNo ?? "",
                    LotNo = LotNo ?? "",
                    IsCCItem = SelectedPlacementItem.IsCCItem,
                    RISD_RefID = SelectedPlacementItem.RISD_EntryID
                };

                await MaterialPlacementService.SavePlacementAsync(request, CurrentUserName, Environment.MachineName);
                
                NotificationService.ShowSuccess("Success", "Material placed successfully.");
                CloseModal();
                await LoadGridDataAsync(); // Refresh list to update balance
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", "Failed to place material: " + ex.Message);
            }
        }
    }
}
