using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services;
using Impulse.Services.Company;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Company
{
    public partial class StoresTree : ComponentBase
    {
        [Inject] private IStoresService StoresService { get; set; } = default!;
        [Inject] private IBlazorContextMenuService ContextMenuService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavManager { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;

        private List<StoreNodeModel> treeData = new();
        private bool isLoading = true;
        private string searchText = "";

        // Modal State
        private bool showModal = false;
        private bool isEditing = false;
        private string modalTitle = "";
        private string modalInputText = "";
        
        // Tracking context for save
        private NodeType currentNodeType;
        private StoreNodeModel? selectedStore;
        private RackNodeModel? selectedRack;
        private BinNodeModel? selectedBin;

        private enum NodeType { None, Store, Rack, Bin }

        private IEnumerable<StoreNodeModel> FilteredTreeData
        {
            get
            {
                if (string.IsNullOrWhiteSpace(searchText))
                    return treeData;

                var q = searchText.Trim().ToLower();
                var result = new List<StoreNodeModel>();

                foreach (var store in treeData)
                {
                    bool storeMatch = store.StoreName.ToLower().Contains(q);
                    
                    var matchingRacks = new List<RackNodeModel>();
                    foreach (var rack in store.Racks)
                    {
                        bool rackMatch = rack.RackNo.ToLower().Contains(q);
                        
                        var matchingBins = new List<BinNodeModel>();
                        foreach (var bin in rack.Bins)
                        {
                            if (storeMatch || rackMatch || bin.ShelfNo.ToLower().Contains(q))
                            {
                                matchingBins.Add(bin);
                            }
                        }

                        if (storeMatch || rackMatch || matchingBins.Any())
                        {
                            var rackClone = new RackNodeModel
                            {
                                RackID = rack.RackID,
                                RackNo = rack.RackNo,
                                StoreID = rack.StoreID,
                                IsExpanded = true,
                                Bins = (storeMatch || rackMatch) ? rack.Bins : matchingBins
                            };
                            matchingRacks.Add(rackClone);
                        }
                    }

                    if (storeMatch || matchingRacks.Any())
                    {
                        var storeClone = new StoreNodeModel
                        {
                            StoreID = store.StoreID,
                            StoreName = store.StoreName,
                            IsExpanded = true,
                            Racks = storeMatch ? store.Racks : matchingRacks
                        };
                        result.Add(storeClone);
                    }
                }
                return result;
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            isLoading = true;
            try
            {
                var oldState = new HashSet<string>();
                foreach (var s in treeData)
                {
                    if (s.IsExpanded) oldState.Add("S_" + s.StoreID);
                    foreach (var rack in s.Racks)
                    {
                        if (rack.IsExpanded) oldState.Add("R_" + rack.RackID);
                    }
                }

                treeData = await StoresService.GetFullTreeAsync();

                if (oldState.Count > 0)
                {
                    foreach (var s in treeData)
                    {
                        s.IsExpanded = oldState.Contains("S_" + s.StoreID);
                        foreach (var rack in s.Racks)
                        {
                            rack.IsExpanded = oldState.Contains("R_" + rack.RackID);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load data: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task RefreshTree()
        {
            await LoadDataAsync();
            NotificationService.Notify(NotificationSeverity.Info, "Refreshed", "Tree data refreshed.");
        }

        private void ToggleStore(StoreNodeModel store)
        {
            store.IsExpanded = !store.IsExpanded;
        }

        private void ToggleRack(RackNodeModel rack)
        {
            rack.IsExpanded = !rack.IsExpanded;
        }

        private void ExpandAll()
        {
            foreach (var s in treeData)
            {
                s.IsExpanded = true;
                foreach (var rack in s.Racks)
                {
                    rack.IsExpanded = true;
                }
            }
        }

        private void CollapseAll()
        {
            foreach (var s in treeData)
            {
                s.IsExpanded = false;
                foreach (var rack in s.Racks)
                {
                    rack.IsExpanded = false;
                }
            }
        }

        // Context Menus
        private async Task ShowStoreMenu(MouseEventArgs e, StoreNodeModel store)
        {
            await ContextMenuService.ShowMenu("storeContextMenu", (int)e.ClientX, (int)e.ClientY, store);
        }

        private async Task ShowRackMenu(MouseEventArgs e, RackNodeModel rack)
        {
            await ContextMenuService.ShowMenu("rackContextMenu", (int)e.ClientX, (int)e.ClientY, rack);
        }

        private async Task ShowBinMenu(MouseEventArgs e, BinNodeModel bin)
        {
            await ContextMenuService.ShowMenu("binContextMenu", (int)e.ClientX, (int)e.ClientY, bin);
        }

        // Action Handlers
        private void AddNewStore()
        {
            currentNodeType = NodeType.Store;
            isEditing = false;
            modalTitle = "Add New Store";
            modalInputText = "";
            showModal = true;
        }

        private void EditStore(StoreNodeModel store)
        {
            currentNodeType = NodeType.Store;
            isEditing = true;
            selectedStore = store;
            modalTitle = "Rename Store";
            modalInputText = store.StoreName;
            showModal = true;
        }

        private async Task ConfirmDeleteStore(StoreNodeModel store)
        {
            try
            {
                var success = await StoresService.DeleteStoreAsync(store.StoreID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Store deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete store: {ex.Message}");
            }
        }

        private void AddNewRack(StoreNodeModel store)
        {
            currentNodeType = NodeType.Rack;
            isEditing = false;
            selectedStore = store;
            modalTitle = $"Add New Rack to {store.StoreName}";
            modalInputText = "";
            showModal = true;
        }

        private void EditRack(RackNodeModel rack)
        {
            currentNodeType = NodeType.Rack;
            isEditing = true;
            selectedRack = rack;
            modalTitle = "Rename Rack";
            modalInputText = rack.RackNo;
            showModal = true;
        }

        private async Task ConfirmDeleteRack(RackNodeModel rack)
        {
            try
            {
                var success = await StoresService.DeleteRackAsync(rack.RackID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Rack deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete rack: {ex.Message}");
            }
        }

        private void AddNewBin(RackNodeModel rack)
        {
            currentNodeType = NodeType.Bin;
            isEditing = false;
            selectedRack = rack;
            modalTitle = $"Add New Bin to {rack.RackNo}";
            modalInputText = "";
            showModal = true;
        }

        private void EditBin(BinNodeModel bin)
        {
            currentNodeType = NodeType.Bin;
            isEditing = true;
            selectedBin = bin;
            modalTitle = "Rename Bin";
            modalInputText = bin.ShelfNo;
            showModal = true;
        }

        private async Task ConfirmDeleteBin(BinNodeModel bin)
        {
            try
            {
                var success = await StoresService.DeleteBinAsync(bin.BinID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Bin deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete bin: {ex.Message}");
            }
        }

        private void CloseModal()
        {
            showModal = false;
        }

        private async Task SaveRecord()
        {
            if (string.IsNullOrWhiteSpace(modalInputText))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Name cannot be empty.");
                return;
            }

            try
            {
                bool success = false;
                switch (currentNodeType)
                {
                    case NodeType.Store:
                        if (isEditing && selectedStore != null)
                            success = await StoresService.UpdateStoreAsync(selectedStore.StoreID, modalInputText);
                        else
                            success = await StoresService.InsertStoreAsync(modalInputText);
                        break;
                    case NodeType.Rack:
                        if (isEditing && selectedRack != null)
                            success = await StoresService.UpdateRackAsync(selectedRack.RackID, modalInputText);
                        else if (!isEditing && selectedStore != null)
                        {
                            success = await StoresService.InsertRackAsync(selectedStore.StoreID, modalInputText);
                            selectedStore.IsExpanded = true;
                        }
                        break;
                    case NodeType.Bin:
                        if (isEditing && selectedBin != null)
                            success = await StoresService.UpdateBinAsync(selectedBin.BinID, modalInputText);
                        else if (!isEditing && selectedRack != null)
                        {
                            success = await StoresService.InsertBinAsync(selectedRack.RackID, modalInputText);
                            selectedRack.IsExpanded = true;
                        }
                        break;
                }

                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", "Record saved successfully.");
                    showModal = false;
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to save: {ex.Message}");
            }
        }

        private void NavigateToDashboard()
        {
            NavManager.NavigateTo("/company");
        }

        private void PrintReport()
        {
            var reportRequest = new ReportRequest
            {
                ReportName = "Stores_Report.rpt"
            };

            ReportNavigationService.PrintReportAsync(reportRequest);
        }
    }
}
