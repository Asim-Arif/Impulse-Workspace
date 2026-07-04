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
    public partial class SteelTree : ComponentBase
    {
        [Inject] private ISteelService SteelService { get; set; } = default!;
        [Inject] private IBlazorContextMenuService ContextMenuService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavManager { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;

        private List<SteelTypeNodeModel> treeData = new();
        private List<SteelCompModel> allCompositions = new();
        private List<string> availableUnits = new();
        
        private bool isLoading = true;
        private string searchText = "";

        // Modal State
        private bool showSteelModal = false;
        private bool showGageModal = false;
        private bool isEditing = false;
        private string modalTitle = "";
        
        private SteelTypeNodeModel? currentSteelModal;
        private SteelGageNodeModel? currentGageModal;

        private IEnumerable<SteelTypeNodeModel> FilteredTreeData
        {
            get
            {
                if (string.IsNullOrWhiteSpace(searchText))
                    return treeData;

                var q = searchText.Trim().ToLower();
                var result = new List<SteelTypeNodeModel>();

                foreach (var steel in treeData)
                {
                    bool steelMatch = steel.SteelType.ToLower().Contains(q) || steel.Description.ToLower().Contains(q);
                    
                    var matchingGages = new List<SteelGageNodeModel>();
                    foreach (var gage in steel.Gages)
                    {
                        if (steelMatch || gage.Gage.ToLower().Contains(q) || gage.GageUnit.ToLower().Contains(q))
                        {
                            matchingGages.Add(gage);
                        }
                    }

                    if (steelMatch || matchingGages.Any())
                    {
                        var steelClone = new SteelTypeNodeModel
                        {
                            SteelID = steel.SteelID,
                            SteelType = steel.SteelType,
                            Description = steel.Description,
                            IsExpanded = true,
                            Gages = steelMatch ? steel.Gages : matchingGages
                        };
                        result.Add(steelClone);
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
                var oldState = new HashSet<int>();
                foreach (var s in treeData)
                {
                    if (s.IsExpanded) oldState.Add(s.SteelID);
                }

                treeData = await SteelService.GetFullSteelTreeAsync();
                allCompositions = await SteelService.GetAllCompositionsAsync();
                availableUnits = await SteelService.GetUnitsAsync();

                if (oldState.Count > 0)
                {
                    foreach (var s in treeData)
                    {
                        s.IsExpanded = oldState.Contains(s.SteelID);
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

        private void ToggleSteel(SteelTypeNodeModel steel)
        {
            steel.IsExpanded = !steel.IsExpanded;
        }

        private void ExpandAll()
        {
            foreach (var s in treeData) s.IsExpanded = true;
        }

        private void CollapseAll()
        {
            foreach (var s in treeData) s.IsExpanded = false;
        }

        // Context Menus
        private async Task ShowSteelMenu(MouseEventArgs e, SteelTypeNodeModel steel)
        {
            await ContextMenuService.ShowMenu("steelContextMenu", (int)e.ClientX, (int)e.ClientY, steel);
        }

        private async Task ShowGageMenu(MouseEventArgs e, SteelGageNodeModel gage)
        {
            await ContextMenuService.ShowMenu("gageContextMenu", (int)e.ClientX, (int)e.ClientY, gage);
        }

        // --- STEEL TYPE ACTIONS ---
        private void AddNewSteelType()
        {
            isEditing = false;
            modalTitle = "Add New Steel Type";
            
            // Prepare a fresh model with ALL compositions ready to be filled
            currentSteelModal = new SteelTypeNodeModel();
            currentSteelModal.Compositions = allCompositions.Select(c => new SteelCompModel
            {
                CompID = c.CompID,
                Symbol = c.Symbol,
                Name = c.Name,
                MinValue = null,
                MaxValue = null
            }).ToList();

            showSteelModal = true;
        }

        private void EditSteelType(SteelTypeNodeModel steel)
        {
            isEditing = true;
            modalTitle = "Edit Steel Type";
            
            // Clone the steel model so we don't modify the tree until save
            currentSteelModal = new SteelTypeNodeModel
            {
                SteelID = steel.SteelID,
                SteelType = steel.SteelType,
                Description = steel.Description,
                StripeMinW = steel.StripeMinW,
                StripeMaxW = steel.StripeMaxW
            };

            // Merge existing compositions with ALL compositions
            var existingComps = steel.Compositions ?? new List<SteelCompModel>();
            currentSteelModal.Compositions = allCompositions.Select(c => {
                var existing = existingComps.FirstOrDefault(ec => ec.CompID == c.CompID);
                return new SteelCompModel
                {
                    CompID = c.CompID,
                    Symbol = c.Symbol,
                    Name = c.Name,
                    MinValue = existing?.MinValue,
                    MaxValue = existing?.MaxValue
                };
            }).ToList();

            showSteelModal = true;
        }

        private async Task SaveSteelTypeRecord()
        {
            if (string.IsNullOrWhiteSpace(currentSteelModal?.SteelType))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Steel Type name cannot be empty.");
                return;
            }

            try
            {
                bool success = await SteelService.SaveSteelTypeAsync(currentSteelModal);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", "Steel Type saved successfully.");
                    showSteelModal = false;
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to save steel type: {ex.Message}");
            }
        }

        private async Task ConfirmDeleteSteelType(SteelTypeNodeModel steel)
        {
            try
            {
                var success = await SteelService.DeleteSteelTypeAsync(steel.SteelID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Steel Type deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete steel type: {ex.Message}");
            }
        }

        // --- GAGE ACTIONS ---
        private void AddNewGage(SteelTypeNodeModel steel)
        {
            isEditing = false;
            modalTitle = $"Add Gage to {steel.SteelType}";
            currentGageModal = new SteelGageNodeModel { SteelID = steel.SteelID };
            showGageModal = true;
        }

        private void EditGage(SteelGageNodeModel gage)
        {
            isEditing = true;
            modalTitle = "Edit Gage";
            currentGageModal = new SteelGageNodeModel
            {
                GageID = gage.GageID,
                SteelID = gage.SteelID,
                Gage = gage.Gage,
                GageUnit = gage.GageUnit
            };
            showGageModal = true;
        }

        private async Task SaveGageRecord()
        {
            if (string.IsNullOrWhiteSpace(currentGageModal?.Gage))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Gage value cannot be empty.");
                return;
            }

            try
            {
                bool success = await SteelService.SaveSteelGageAsync(currentGageModal, !isEditing);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", "Gage saved successfully.");
                    showGageModal = false;
                    
                    // Keep parent expanded
                    var parent = treeData.FirstOrDefault(s => s.SteelID == currentGageModal.SteelID);
                    if (parent != null) parent.IsExpanded = true;
                    
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to save gage: {ex.Message}");
            }
        }

        private async Task ConfirmDeleteGage(SteelGageNodeModel gage)
        {
            try
            {
                var success = await SteelService.DeleteSteelGageAsync(gage.GageID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Gage deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete gage: {ex.Message}");
            }
        }

        private void CloseModals()
        {
            showSteelModal = false;
            showGageModal = false;
        }

        private void NavigateToDashboard()
        {
            NavManager.NavigateTo("/company");
        }

        private void PrintReport()
        {
            // Legacy printed Steel_Report.rpt but this wasn't specifically mentioned for SteelList.
            // Leaving hook here in case.
            var reportRequest = new ReportRequest
            {
                ReportName = "Steel_Report.rpt"
            };

            ReportNavigationService.PrintReportAsync(reportRequest);
        }
    }
}
