using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.WebUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.VendGateRcvdInspection
{
    public partial class VendGateRcvdInspection : ComponentBase
    {
        [Parameter]
        public string RcvID { get; set; }

        private bool IsLoading = true;
        private string ReturnUrl = "/stock/vend-rcv-list";
        
        private List<VendGateRcvdInspectionMaterialItem> Materials = new List<VendGateRcvdInspectionMaterialItem>();
        private VendGateRcvdInspectionMaterialItem _selectedMaterial;
        public VendGateRcvdInspectionMaterialItem SelectedMaterial
        {
            get => _selectedMaterial;
            set
            {
                if (_selectedMaterial != value)
                {
                    _selectedMaterial = value;
                    _ = LoadInspectionDataAsync();
                }
            }
        }

        private VendGateRcvdInspectionViewModel InspectionModel;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var uri = NavigationManager.ToAbsoluteUri(NavigationManager.Uri);
                if (QueryHelpers.ParseQuery(uri.Query).TryGetValue("returnUrl", out var url))
                {
                    ReturnUrl = url;
                }

                IsLoading = true;
                Materials = await VendGateRcvdInspectionService.GetMaterialsByRcvIDAsync(RcvID);
                
                // Select first material by default if any
                if (Materials.Any())
                {
                    SelectedMaterial = Materials.First();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to load materials: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task<IEnumerable<VendGateRcvdInspectionMaterialItem>> SearchMaterials(string searchText)
        {
            await Task.Delay(0);
            if (string.IsNullOrWhiteSpace(searchText))
                return Materials;
            
            return Materials.Where(m => m.MaterialName.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task LoadInspectionDataAsync()
        {
            if (SelectedMaterial == null)
            {
                InspectionModel = null;
                return;
            }

            try
            {
                InspectionModel = await VendGateRcvdInspectionService.GetInspectionByMaterialAsync(SelectedMaterial.EntryID, SelectedMaterial.GroupID);
                StateHasChanged();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to load inspection data: {ex.Message}");
            }
        }

        private async Task SaveInspection()
        {
            if (InspectionModel == null || SelectedMaterial == null) return;

            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                string userName = authState.User.Identity.Name ?? "Unknown";
                string machineName = "WebClient"; // Usually handled server-side or via JS, placeholder for now.

                await VendGateRcvdInspectionService.SaveInspectionAsync(InspectionModel, userName, machineName);
                
                NotificationService.ShowSuccess("Success", "Inspection saved successfully.");
                
                // Reload to get fresh EntryID in case it was a new insert
                await LoadInspectionDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to save inspection: {ex.Message}");
            }
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo(ReturnUrl);
        }
    }
}
